"""Устанавливает принятый VBA-пакет без изменения данных и архивирует первичные доказательства."""
import hashlib
import json
import os
from pathlib import Path
from shutil import copyfile
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = next(p for p in Path(__file__).resolve().parents if (p / 'AGENTS.md').is_file())
FULL = Path(__file__).resolve().parent.parent
TYPED = ROOT / 'docs/regression/NormativeCalculators_2026-11-24/TorsionTyped_2026-11-25'


def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


review = read(FULL / 'Review/FinalReview.json')
proof = read(TYPED / 'PackageProof.json')
assert review['Status'] == read(FULL / 'GoalAcceptance.json')['Status'] == 'PASS'
assert read(TYPED / 'DirectedAcceptance.json')['Status'] == 'PASS'
book = ROOT / 'workbook/output/RC_Section_NDM.xlsm'
export = ROOT / 'workbook/output/VBA_All_Code.txt'
assert sha(book) == proof['BaselineSHA256']
assert sha(TYPED / 'Candidate.xlsm') == proof['CandidateSHA256']
assert sha(TYPED / 'VBA_All_Code.txt') == proof['ExportSHA256']
with ZipFile(book) as previous, ZipFile(TYPED / 'Candidate.xlsm') as candidate:
    assert previous.namelist() == candidate.namelist()
    changed = [n for n in previous.namelist() if previous.read(n) != candidate.read(n)]
    assert changed == ['xl/vbaProject.bin']
with ZipFile(TYPED / 'Candidate.xlsm') as candidate, ZipFile(TYPED / 'CandidateBuild02.xlsm') as compiled:
    assert candidate.read('xl/vbaProject.bin') == compiled.read('xl/vbaProject.bin')
foreign = [ROOT / 'workbook/output/NDM_Test_1232_30/RC_Section_NDM_execution_report.txt',
           ROOT / 'workbook/output/RC_Section_NDM_execution_report.txt', ROOT / 'Типы_сечения.dwg']
foreign_before = {p.relative_to(ROOT).as_posix(): sha(p) if p.exists() else None for p in foreign}
for source, destination in ((TYPED / 'Candidate.xlsm', book), (TYPED / 'VBA_All_Code.txt', export)):
    temporary = destination.with_name(destination.name + '.accepted.tmp')
    assert not temporary.exists()
    copyfile(source, temporary)
    assert sha(source) == sha(temporary)
    os.replace(temporary, destination)
assert sha(book) == proof['CandidateSHA256'] and sha(export) == proof['ExportSHA256']
assert foreign_before == {p.relative_to(ROOT).as_posix(): sha(p) if p.exists() else None for p in foreign}
installation = {'Status': 'PASS', 'InstalledBookSHA256': sha(book), 'InstalledExportSHA256': sha(export),
                'OnlyChangedZIPParts': changed, 'AllNonVBAPartsPreserved': True,
                'NativeCompiledVBAInstalled': True, 'UnrelatedFilesPreserved': foreign_before,
                'FullAcceptanceSource': proof['BaselineSHA256'],
                'TypedChangeAcceptance': 'DirectedAcceptance.json:40 executions; no new full run'}
(FULL / 'Review/Installation.json').write_text(json.dumps(installation, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')


def archive(directory, filename):
    target = directory / filename
    assert not target.exists()
    included, excluded = [], []
    paths = sorted(p for p in directory.rglob('*') if p.is_file() and p != target)
    with ZipFile(target, 'w', compression=ZIP_DEFLATED, compresslevel=9) as output:
        for path in paths:
            relative = path.relative_to(directory).as_posix()
            record = {'Path': relative, 'SHA256': sha(path), 'Bytes': path.stat().st_size}
            if path.suffix.lower() in {'.xlsm', '.xlsx', '.xls', '.dwg', '.bak', '.zip', '.pyc'}:
                excluded.append(record)
            else:
                output.write(path, relative)
                included.append(record)
    with ZipFile(target) as check:
        assert check.testzip() is None
        assert len(check.namelist()) == len(included)
        for item in included:
            assert hashlib.sha256(check.read(item['Path'])).hexdigest().upper() == item['SHA256']
    inventory = {'Archive': filename, 'SHA256': sha(target), 'Included': included,
                 'ExcludedLocalBinaryArtifacts': excluded, 'OriginalFilesRetainedLocally': True}
    (directory / 'Review').mkdir(exist_ok=True)
    (directory / 'Review/EvidenceInventory.json').write_text(json.dumps(inventory, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    return {'Archive': str(target.relative_to(ROOT)), 'SHA256': sha(target), 'Files': len(included), 'Bytes': target.stat().st_size}


archives = [archive(FULL, 'FullEvidence.zip'), archive(TYPED, 'TypedEvidence.zip')]
print(json.dumps({'Installation': installation, 'Archives': archives}, ensure_ascii=False))
