"""Проверяет направленную приёмку диагностики исчезнувших участков и итоговую книгу."""
import hashlib
import json
from pathlib import Path
import re
import zipfile

ROOT = Path(__file__).resolve().parents[2]
EVIDENCE = ROOT / 'docs/regression/CollapsedGeometryMessages_2026-10-19'
SUITES = {
    'FinalChecks.txt': [
        'modTestGeometryReview.RunCollapsedInputMessageTests',
        'modTestGeometryReview.RunCollapsedInputAddressTests',
        'modTestDeclarativeGeometry.RunCoreGeometryTests',
        'modTestGeometryReview.RunGeometryReviewTests',
        'modTestSectionRecipes.RunCircleRecipeInputTests',
        'modTestSectionRecipes.RunCircleBindingTests',
        'modTestSectionRecipes.RunRecipeCaptionTests',
        'modTestRecipeOperations.RunRecipeDiagnosticTests',
    ],
    'FinalGeometryControls.txt': [
        'modTestGeometry.RunGeometryTests',
        'modTestGeometryReview.RunRoundedEndpointTests',
        'modTestSectionRecipes.RunRecipeRuleTests',
        'modTestRecipeOperations.RunRecipeOperationTests',
    ],
}


def digest(data):
    return hashlib.sha256(data).hexdigest().upper()


def receipt(name):
    return json.loads((EVIDENCE / name).read_text(encoding='utf-8-sig'))


def inspect_logs(archive, candidate_sha):
    count = 0
    with zipfile.ZipFile(archive) as logs:
        for name, expected in SUITES.items():
            text = logs.read(name).decode('utf-8-sig').replace('\r\n', '\n')
            assert f'SOURCE_SHA256: {candidate_sha}' in text, name
            assert 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' in text, name
            assert not re.search(r'^FAIL:|^SCRIPT ERROR:|^RUNTIME ERROR|failed=[1-9]', text, re.M), name
            found = re.findall(r'^===== (\w+\.\w+) =====$', text, re.M)
            assert found == expected, (name, found)
            for macro in expected:
                assert f'SUITE_FINISHED: {macro};' in text, macro
            totals = re.findall(r'^TOTAL[^\n]*?passed=(\d+); failed=(\d+)', text, re.M)
            assert totals and all(int(failed) == 0 for _, failed in totals), name
            count += sum(int(passed) for passed, _ in totals)
    return count


def verify():
    acceptance = receipt('Acceptance.json')
    artifact = receipt('PublishedArtifact.json')
    equality = receipt('PublishedEquality.json')
    package = receipt('Package.json')
    coverage = receipt('Coverage.json')
    assert acceptance['Passed'] and artifact['Passed'] and equality['Passed']
    assert not acceptance['FullOnOffPerformed']
    assert acceptance['AcceptedSuites'] == sum(map(len, SUITES.values())) == 12
    count = inspect_logs(EVIDENCE / 'Logs.zip', package['AcceptedSHA256'])
    assert count == acceptance['AssertionsPassed'] and acceptance['AssertionsFailed'] == 0
    sources = [p for folder in ('src', 'tests') for p in (ROOT / folder).rglob('*') if p.suffix in ('.cls', '.bas')]
    assert len(sources) == equality['SourceComponents'] == 155 and equality['ExportComponents'] == 160
    assert not equality['Missing'] and not equality['Differences'] and not equality['UnexpectedCodeComponents']
    assert len(artifact['SourceTextSHA256']) == len(sources)
    for path in sources:
        assert digest(path.read_text(encoding='utf-8-sig').encode('utf-8')) == artifact['SourceTextSHA256'][path.stem], path
    assert coverage['ProductionComponents'] == 112 and coverage['ProductionClasses'] == 97
    assert len(coverage['Files']) == 112
    comments = receipt('ClassComments.json')
    assert comments['status'] == 'PASS' and comments['missingFieldCount'] == 0
    output = ROOT / 'workbook/output'
    book = (output / 'RC_Section_NDM.xlsm').read_bytes()
    assert digest(book) == artifact['WorkbookSHA256'] == package['PublishedCandidateSHA256']
    assert digest((output / 'VBA_All_Code.txt').read_bytes()) == artifact['AllCodeSHA256'] == artifact['NativeExportSHA256']
    assert package['ChangedParts'] == ['xl/vbaProject.bin']
    with zipfile.ZipFile(output / 'RC_Section_NDM.xlsm') as workbook:
        assert set(workbook.namelist()) == {p['Part'] for p in package['UnchangedParts']} | {'xl/vbaProject.bin'}
        for part in package['UnchangedParts']:
            assert digest(workbook.read(part['Part'])) == part['SHA256'], part['Part']
        assert digest(workbook.read('xl/vbaProject.bin')) == artifact['VbaProjectSHA256']
    return dict(Passed=True, Suites=12, Assertions=count, PreservedWorkbookParts=len(package['UnchangedParts']),
                SourceComponents=len(sources), WorkbookSHA256=artifact['WorkbookSHA256'])


if __name__ == '__main__':
    print(json.dumps(verify(), ensure_ascii=False))
