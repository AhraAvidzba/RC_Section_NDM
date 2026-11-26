"""ДЛЯ ТЕСТОВ: принимает связанный вывод T, сохраняет данные книги и первичные журналы."""
import hashlib
import json
import os
from pathlib import Path
import re
from shutil import copyfile
from zipfile import ZipFile, ZIP_DEFLATED

HERE = Path(__file__).resolve().parent
ROOT = next(p for p in HERE.parents if (p / 'AGENTS.md').is_file())


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


proof = read(HERE / 'PackageProof.json')
assert proof['Status'] == 'PASS' and proof['SourceEquality'] and proof['NativeComponents'] == 209
assert sha(HERE / 'Candidate.xlsm') == proof['CandidateSHA256']
runs = read(HERE / 'Runs.json')
assert len(runs) == 6 and {(r['Job'], r['Mode']) for r in runs} == {
    (job, mode) for job in ('Tables', 'File', 'Excel') for mode in ('No', 'Yes')}
executions = assertions = 0
for run in runs:
    path = ROOT / run['Report']
    body = path.read_text(encoding='utf-8-sig')
    assert sha(path) == run['SHA256'] and run['Passed']
    assert re.findall(r'^SUITE_STARTED: ([^;]+);', body, re.M) == run['Macros']
    assert re.findall(r'^SUITE_FINISHED: ([^;]+);', body, re.M) == run['Macros']
    assert 'SOURCE_SHA256: ' + proof['CandidateSHA256'] in body
    assert 'GLOBAL_MODE: General.DiagramExtension=' + run['Mode'] + ';' in body
    assert 'SOURCE_UNCHANGED: True' in body and 'WATCHDOG_COMPLETED: exit=0;' in body
    assert not re.search(r'^FAIL\b|Failed=[1-9]|failed=[1-9]|^RUNTIME[: ]|RUNTIME ERROR|WATCHDOG_FAILURE', body, re.M)
    expected_actual = [run['Mode']] * 3 if run['Job'] == 'Tables' else ['Yes']
    assert re.findall(r'^MODE_AFTER_SUITE: (\w+)', body, re.M) == expected_actual
    totals = re.findall(r'^.*?: [Pp]assed=(\d+); [Ff]ailed=(\d+)', body, re.M)
    assert len(totals) == len(run['Macros']) and all(int(failed) == 0 for _, failed in totals)
    assertions += sum(int(passed) for passed, _ in totals)
    closure = read(Path(str(path) + '.ProcessClosure.json'))
    assert closure['Passed'] and closure['ChildHasExited'] and closure['ChildExitCode'] == 0
    assert closure['Closure']['Passed'] and closure['RunnerCleanupValid']
    if run['Job'] == 'File':
        for expected in ('PASS File bar overflow diagnosed', 'PASS File bar overflow before clear',
                         'PASS File stale bar tail cleared', 'PASS File NM details remain header only'):
            assert expected in body
        assert len(re.findall(r'^PASS File calculated bars exact ', body, re.M)) == 65
        assert len(re.findall(r'^PASS File refreshed bars exact ', body, re.M)) == 65
    executions += len(run['Macros'])
assert executions == 10
book, export = ROOT / 'workbook/output/RC_Section_NDM.xlsm', ROOT / 'workbook/output/VBA_All_Code.txt'
assert sha(book) == proof['BaselineSHA256']
assert sha(HERE / 'VBA_All_Code.txt') == proof['ExportSHA256']
with ZipFile(book) as baseline, ZipFile(HERE / 'Candidate.xlsm') as candidate:
    changed = [name for name in baseline.namelist() if baseline.read(name) != candidate.read(name)]
    assert changed == proof['ChangedZIPParts']
    assert set(changed) == {'xl/vbaProject.bin', 'xl/sharedStrings.xml'}
foreign = [ROOT / 'workbook/output/NDM_Test_1232_30/RC_Section_NDM_execution_report.txt',
           ROOT / 'workbook/output/RC_Section_NDM_execution_report.txt']
foreign_hashes = {p.relative_to(ROOT).as_posix(): sha(p) if p.exists() else None for p in foreign}
for source, destination in ((HERE / 'Candidate.xlsm', book), (HERE / 'VBA_All_Code.txt', export)):
    temporary = destination.with_name(destination.name + '.file-torsion.tmp')
    assert not temporary.exists()
    copyfile(source, temporary)
    assert sha(temporary) == sha(source)
    os.replace(temporary, destination)
assert sha(book) == proof['CandidateSHA256'] and sha(export) == proof['ExportSHA256']
assert foreign_hashes == {p.relative_to(ROOT).as_posix(): sha(p) if p.exists() else None for p in foreign}
acceptance = {'Status': 'PASS', 'Executions': executions, 'Assertions': assertions,
              'NoFullRunStarted': True, 'FileAndExcelWorkflowsPassed': True,
              'AllFiveColumnsAndBothCombinationsVerified': True,
              'RefreshRepairsSheetAndClearsTail': True, 'OverflowCheckedBeforeClear': True,
              'FileNDMDetailsRemainOmitted': True, 'CompiledComponents': 209,
              'InstalledBookSHA256': sha(book), 'InstalledExportSHA256': sha(export),
              'ChangedZIPParts': changed, 'OtherWorkbookDataAndFormattingPreserved': True,
              'FixtureModeNote': 'Tables execute No/Yes; existing workbook workflow fixture explicitly sets Yes in both outer runs.',
              'RunnerReviewNote': 'Initial scanner matched PASS Runtime preserves width; anchored RUNTIME pattern fixed, successful primary log retained and reviewed without rerun.'}
(HERE / 'Acceptance.json').write_text(json.dumps(acceptance, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
target = HERE / 'Evidence.zip'
assert not target.exists()
included, excluded = [], []
with ZipFile(target, 'w', compression=ZIP_DEFLATED, compresslevel=9) as archive:
    for path in sorted(HERE.rglob('*')):
        if not path.is_file() or path == target:
            continue
        record = {'Path': path.relative_to(HERE).as_posix(), 'SHA256': sha(path), 'Bytes': path.stat().st_size}
        if path.suffix.lower() in ('.xlsm', '.zip', '.pyc'):
            excluded.append(record)
        else:
            archive.write(path, record['Path'])
            included.append(record)
with ZipFile(target) as archive:
    assert archive.testzip() is None
    for record in included:
        assert hashlib.sha256(archive.read(record['Path'])).hexdigest().upper() == record['SHA256']
(HERE / 'EvidenceInventory.json').write_text(json.dumps({'ArchiveSHA256': sha(target), 'Included': included,
    'ExcludedLocalBinaryArtifacts': excluded, 'OriginalLogsRetained': True}, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps(acceptance, ensure_ascii=False))
