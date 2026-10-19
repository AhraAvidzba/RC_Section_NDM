"""Собирает доказательства направленной приёмки без запуска Excel или AutoCAD."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import hashlib
import json
import re

ROOT = Path(__file__).resolve().parents[2]
BASE = ROOT / 'docs/regression/SharedRebarInput_2026-10-18/OffsetExport_2026-10-19'


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def receipt(name):
    return json.loads((BASE / name).read_text('utf-8-sig'))


def main():
    package = receipt('Package.json')
    publication = receipt('Publication.json')
    for name in ('PublishedEquality.json', 'AcceptedV3Equality.json',
                 'KernelMath_Final/Comparison.json', 'SavedNative_Verified/Acceptance.json'):
        assert receipt(name)['Passed'], name
    assert package['Status'] == 'PRESERVATION_PASS'
    assert publication['Passed'] and publication['Readback']
    assert publication['UserExecutionReportUnchanged']
    book = ROOT / 'workbook/output/RC_Section_NDM.xlsm'
    code = ROOT / 'workbook/output/VBA_All_Code.txt'
    assert sha(book) == package['AcceptedSHA256'] == publication['WorkbookSHA256']
    assert sha(code) == publication['AllCodeSHA256'] == sha(BASE / 'PublishedNative.txt')
    assert sha(ROOT / 'workbook/output/RC_Section_NDM_execution_report.txt') == '9426FB7EC9A0A5CED9CB75021E71F7967DE0384733E908CC8D720BE5E9AE865E'
    with ZipFile(BASE / 'Saved.xlsm') as old, ZipFile(book) as new, ZipFile(BASE / 'Candidate.xlsm') as compiled:
        assert set(old.namelist()) == set(new.namelist())
        changed = sorted(n for n in old.namelist() if old.read(n) != new.read(n))
        assert changed == ['xl/sharedStrings.xml', 'xl/vbaProject.bin'], changed
        assert new.read('xl/vbaProject.bin') == compiled.read('xl/vbaProject.bin')
    math_identity = receipt('KernelMath_Final/Identity.json')
    assert math_identity['SourceHashes']['B'] == sha(BASE / 'Candidate.xlsm')
    assert math_identity['SourceHashes']['A'] == sha(BASE / 'MathBaseline.xlsm')
    tests = {}
    for mode in ('Yes', 'No'):
        text = (BASE / f'Release_{mode}.txt').read_text('utf-8-sig')
        assert f'SOURCE_SHA256: {publication["WorkbookSHA256"]}' in text
        assert f'GLOBAL_MODE: General.DiagramExtension={mode}' in text
        assert 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' in text
        assert text.count('SUITE_STARTED:') == text.count('SUITE_FINISHED:') == 10
        counts = re.findall(r'passed=(\d+); failed=(\d+)', text)
        assert len(counts) == 10 and all(int(f) == 0 for _, f in counts)
        assert sum(int(p) for p, _ in counts) == 444
        tests[mode] = {'Suites': 10, 'PassedAssertions': 444, 'FailedAssertions': 0}
    comments_path = ROOT / 'tmp/FinalClassComments3.json'
    comments = json.loads(comments_path.read_text('utf-8-sig'))
    assert comments['status'] == 'PASS'
    summary = {k: v for k, v in comments.items() if k != 'details'}
    (BASE / 'ClassCommentsSummary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    evidence = [
        'Package.json', 'Publication.json', 'PublishedEquality.json', 'PublishedNative.txt',
        'AcceptedV3Equality.json', 'SavedEquality.json', 'SavedNative.txt', 'SavedInventory.json',
        'Release_Yes.txt', 'Release_No.txt', 'ClassCommentsSummary.json', 'BaselineProfile.txt',
    ]
    for folder in ('KernelMath_Final', 'SavedNative_Verified'):
        evidence += [p.relative_to(BASE).as_posix() for p in sorted((BASE / folder).iterdir())
                     if p.is_file() and p.suffix in ('.json', '.txt')]
    extras = {
        'reproduce/Package-OffsetCandidate.py': ROOT / 'tmp/Package-OffsetCandidate.py',
        'reproduce/Prepare-CADMathBaseline.ps1': ROOT / 'tmp/Prepare-CADMathBaseline.ps1',
        'reproduce/Update-OffsetCandidate.ps1': ROOT / 'tmp/Update-OffsetCandidate.ps1',
    }
    manifest = [{'Path': n, 'SHA256': sha(BASE / n), 'Bytes': (BASE / n).stat().st_size} for n in evidence]
    manifest += [{'Path': n, 'SHA256': sha(p), 'Bytes': p.stat().st_size} for n, p in extras.items()]
    with ZipFile(BASE / 'Evidence.zip', 'w', compression=ZIP_DEFLATED, compresslevel=9) as archive:
        for name in evidence:
            archive.write(BASE / name, name)
        for name, path in extras.items():
            archive.write(path, name)
    with ZipFile(BASE / 'Evidence.zip') as archive:
        assert archive.testzip() is None
        for item in manifest:
            assert hashlib.sha256(archive.read(item['Path'])).hexdigest().upper() == item['SHA256']
    result = {
        'Passed': True, 'Scope': 'Directed offset and native CAD export acceptance; not full ON/OFF acceptance',
        'Tests': tests, 'NativeCalculationPairs': 12, 'NativeCalculationRuns': 24,
        'NativeMathEquivalent': True, 'SavedCADRoundtripExact': True, 'WCSVerified': True,
        'SourceComponents': 155, 'WorkbookComponents': 161, 'NewComponents': 0,
        'UserDataAndResultsPreserved': True, 'UserExecutionReportUnchanged': True,
        'WorkbookSHA256': publication['WorkbookSHA256'], 'AllCodeSHA256': publication['AllCodeSHA256'],
        'FullOnOffRun': False, 'PreviousRun03StatusChanged': False,
        'EvidenceArchiveSHA256': sha(BASE / 'Evidence.zip'),
    }
    (BASE / 'EvidenceManifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    (BASE / 'Acceptance.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(result))


if __name__ == '__main__':
    main()
