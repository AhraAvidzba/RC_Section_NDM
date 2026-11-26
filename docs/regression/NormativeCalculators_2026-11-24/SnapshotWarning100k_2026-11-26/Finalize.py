"""ДЛЯ ТЕСТОВ: принимает порог 100000 и устанавливает книгу с сохранением остальных данных."""
import hashlib
import json
import os
from pathlib import Path
import re
from shutil import copyfile
from zipfile import ZipFile, ZIP_DEFLATED

HERE = Path(__file__).resolve().parent
ROOT = next(p for p in HERE.parents if (p / 'AGENTS.md').is_file())
def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))
def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()

proof = read(HERE / 'PackageProof.json')
assert proof['Status'] == 'PASS' and proof['SourceEquality'] and proof['NativeComponents'] == 209
log = (HERE / 'Policy.txt').read_text(encoding='utf-8-sig')
macro = 'modTestCalculationSnapshot.RunSnapshotOutputPolicyTests'
assert re.findall(r'^SUITE_FINISHED: ([^;]+);', log, re.M) == [macro]
assert 'TOTAL_SNAPSHOT_OUTPUT_POLICY: passed=68; failed=0' in log
assert 'SOURCE_SHA256: ' + proof['CandidateSHA256'] in log
assert 'SOURCE_UNCHANGED: True' in log and 'WATCHDOG_COMPLETED: exit=0;' in log
assert not re.search(r'^FAIL\b|failed=[1-9]|^RUNTIME\b|WATCHDOG_FAILURE', log, re.M)
assert read(HERE / 'Policy.txt.ProcessClosure.json')['Passed']
for mesh in ('Adaptive', 'Contour'):
    for mode in ('Excel', 'File'):
        for suffix in ('formerThresholdAbove', 'thresholdBelow', 'thresholdExact', 'stateFile',
                       'boundaryExact', 'boundaryLargeFile', 'mandatoryExact', 'mandatory', 'geometryOverflowNoFile'):
            assert f'PASS rows.policy.{mesh}.{mode}.{suffix}' in log
book, export = ROOT / 'workbook/output/RC_Section_NDM.xlsm', ROOT / 'workbook/output/VBA_All_Code.txt'
assert sha(book) == proof['BaselineSHA256']
for source, destination, digest in ((HERE / 'Candidate.xlsm', book, proof['CandidateSHA256']),
                                    (HERE / 'VBA_All_Code.txt', export, proof['ExportSHA256'])):
    assert sha(source) == digest
    temporary = destination.with_name(destination.name + '.threshold.tmp')
    copyfile(source, temporary)
    assert sha(temporary) == digest
    os.replace(temporary, destination)
acceptance = {'Status': 'PASS', 'Assertions': 68, 'SolverCalled': False, 'FullSuiteStarted': False,
              'Threshold': 100000, 'Comparison': 'strictly greater',
              'Meshes': ['Adaptive', 'Contour'], 'Storages': ['Excel', 'File'],
              'InstalledWorkbookSHA256': sha(book), 'InstalledExportSHA256': sha(export),
              'ChangedZIPParts': proof['ChangedZIPParts'], 'OtherWorkbookDataPreserved': True,
              'NativeLiteralFormatting': 'VBE omits optional # in fractional Double 123.456; same in baseline, numeric value unchanged.'}
(HERE / 'Acceptance.json').write_text(json.dumps(acceptance, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
inventory = []
with ZipFile(HERE / 'Evidence.zip', 'w', compression=ZIP_DEFLATED) as archive:
    for path in sorted(HERE.iterdir()):
        if not path.is_file() or path.suffix in ('.xlsm', '.zip'):
            continue
        included = path.name not in ('BaselineNative.txt', 'NativeCandidate.txt', 'VBA_All_Code.txt')
        inventory.append({'File': path.name, 'SHA256': sha(path), 'Archived': included})
        if included:
            archive.write(path, path.name)
(HERE / 'EvidenceInventory.json').write_text(json.dumps({'ArchiveSHA256': sha(HERE / 'Evidence.zip'),
    'Files': inventory, 'OriginalsRetainedLocally': True}, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps(acceptance, ensure_ascii=False))
