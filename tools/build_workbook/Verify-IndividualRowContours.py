"""Проверяет направленную приёмку индивидуальных контуров R2/R3 и сохранность книги."""
import argparse
import hashlib
import json
from pathlib import Path
import re
import subprocess
import zipfile

ROOT = Path(__file__).resolve().parents[2]
EVIDENCE = ROOT / 'docs/regression/R2R3IndividualContours_2026-10-20'
FINAL = EVIDENCE / 'Final'
BASELINE = 'e75ba9b728193bd4bdb379e8d576b2659fcb7558'
EXPECTED = [
    'modTestIndividualContours.RunOffsetNeighborMergeTests',
    'modTestIndividualContours.RunUserRoundedContourTests',
    'modTestIndividualContours.RunIndividualContourTests',
    'modTestIndividualContours.RunIndividualContourFailureTests',
    'modTestIndividualContours.RunIndividualContourWorkbookTests',
    'modTestAudit02Api.RunRebarContactBuilderTests',
    'modTestAudit02Api.RunIndependentRowOffsetBuilderTests',
    'modTestAudit02Api.RunIndependentRowOffsetGeometryTests',
    'modTestRecipeOperations.RunRecipeOperationTests',
    'modTestRecipeOperations.RunCollapsingOffsetTests',
    'modTestRecipeOperations.RunFullBoundaryOffsetTests',
    'modTestAudit02Api.RunOffsetOpeningAnnotationTests',
    'modTestAudit02Api.RunCircleSetbackTests',
    'modTestAudit02Api.RunOffsetWallCollisionTests',
    'modTestAudit02Api.RunIndependentRowOffsetWorkbookTests',
    'modTestGeometry.RunGeometryTests',
    'modTestDeclarativeGeometry.RunCoreGeometryTests',
    'modTestPolygonSpec.RunPolygonSpecTests',
    'modTestSectionRecipes.RunRecipeRuleTests',
    'modTestZeroContourParity.TestDumpZeroContourParity',
]
SAVED = ['modTestRecipeOperations.RunSavedPointedOffsetTests',
         'modTestRecipeOperations.RunSavedOffsetCalculationTests']
CORRECTED = ['modTestWorkbookInterface.RunCorrectedUiExpectationsTests',
             'modTestGeometryReview.RunCollapsedInputMessageTests',
             'modTestIndividualContours.RunIndividualContourTests',
             'modTestIndividualContours.RunUserRoundedContourTests']


def digest(data):
    return hashlib.sha256(data).hexdigest().upper()


def receipt(name):
    return json.loads((FINAL / name).read_text(encoding='utf-8-sig'))


def zero_payload(text):
    lines = [line for line in text.splitlines() if line.startswith(('ZERO_CASE:', 'C|', 'R|', 'ZERO_PARITY_CASES:'))]
    assert sum(line.startswith('ZERO_CASE:') for line in lines) == 20
    assert lines[-1] == 'ZERO_PARITY_CASES: 20'
    assert sum(line.startswith('C|') for line in lines) == 7832
    assert sum(line.startswith('R|') for line in lines) == 560
    return '\n'.join(lines).encode('utf-8')


def inspect_log(path, expected):
    text = path.read_text(encoding='utf-8-sig')
    assert 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' in text, path
    assert not re.search(r'^FAIL:|^SCRIPT ERROR:|^RUNTIME ERROR|SUITE_ERROR|WATCHDOG_FAILURE|failed=[1-9]', text, re.M), path
    macros = re.findall(r'^===== (\w+\.\w+) =====', text, re.M)
    assert macros == expected, (path, macros)
    for macro in macros:
        assert f'SUITE_FINISHED: {macro};' in text, (path, macro)
    totals = re.findall(r'^TOTAL[^\n]*?passed=(\d+); failed=(\d+)', text, re.M)
    assert totals and all(int(failed) == 0 for _, failed in totals), path
    candidate = re.search(r'^SOURCE_SHA256: ([0-9A-F]{64})$', text, re.M)[1]
    return text, candidate, sum(int(passed) for passed, _ in totals)


def verify(published):
    with zipfile.ZipFile(FINAL / 'BaselineAndDevelopment.zip') as baseline:
        old = baseline.read('ZeroBaseline/ZeroBits.txt').decode('utf-8-sig')
    old_payload = zero_payload(old)
    counts = {}
    hashes = set()
    log_hashes = {}
    for mode in ('Off', 'On'):
        path = FINAL / f'Directed_{mode}.txt'
        text, candidate, count = inspect_log(path, EXPECTED)
        assert zero_payload(text) == old_payload, f'Zero-offset bitwise parity: {mode}'
        hashes.add(candidate)
        counts[mode] = count
        for case in range(8):
            assert f'RECIPE_CASE: {case};' in text, (mode, case)
        for reverse in ('False', 'True'):
            assert f'USER_CASE: reverse={reverse}; top.R2=5; R=40; connectors=40' in text
        assert 'TOTAL_INDIVIDUAL_CONTOURS: passed=269; failed=0' in text
        assert 'DOMAIN_SCOPE: outers=2; openings=2; rows=2,3; independent=True' in text
        assert 'TOTAL_OFFSET_NEIGHBOR_MERGE: passed=264; failed=0' in text
        assert 'TOTAL_INDIVIDUAL_CONTOUR_FAILURES: passed=11; failed=0' in text
        assert 'TOTAL_INDIVIDUAL_CONTOUR_WORKBOOK: passed=604; failed=0' in text
        log_hashes[path.name] = digest(path.read_bytes())
        path = FINAL / f'Saved_{mode}.txt'
        _, candidate, count = inspect_log(path, SAVED)
        hashes.add(candidate)
        counts[f'Saved{mode}'] = count
        log_hashes[path.name] = digest(path.read_bytes())
        path = FINAL / f'Corrected_{mode}.txt'
        corrected, candidate, count = inspect_log(path, CORRECTED)
        assert 'TOTAL_CORRECTED_UI_EXPECTATIONS: passed=121; failed=0' in corrected
        assert 'TOTAL_COLLAPSED_INPUT_MESSAGES: passed=65; failed=0' in corrected
        hashes.add(candidate)
        counts[f'Corrected{mode}'] = count
        log_hashes[path.name] = digest(path.read_bytes())
    assert len(hashes) == 1, hashes
    # Общая проверка не подменяет и не редактирует прежние операции offset.
    for name in ('src/Geometry/CGeometryPath.cls', 'src/Section/CSectionGeometryQuery.cls'):
        old_source = subprocess.check_output(['git', 'show', BASELINE + ':' + name], cwd=ROOT)
        current = (ROOT / name).read_text(encoding='utf-8-sig')
        if name.endswith('CSectionGeometryQuery.cls'):
            # Единственное дополнение — вход в прежний ContainsRaw без повторного копирования массива.
            current, removed = re.subn(r"(?m)^' Проверяет точку по заранее подготовленным границам;.*\n' В отличие.*\nFriend Function ContainsPreparedPoint\([\s\S]*?End Function\n\n", '', current, count=1)
            assert removed == 1
        assert old_source.decode('utf-8-sig').replace('\r\n', '\n') == current, name
    audit = receipt('ClassComments.json')
    assert audit['status'] == 'PASS' and audit['missingFieldCount'] == 0
    result = dict(Passed=True, CandidateSHA256=hashes.pop(), DirectedSuites=2 * (len(EXPECTED) + len(SAVED) + len(CORRECTED)),
                  AssertionsByMode=counts, AssertionsPassed=sum(counts.values()), AssertionsFailed=0,
                  FullOnOffPerformed=False, ZeroOffsetCases=20, ConcreteElementsCompared=7832,
                  RebarsCompared=560, BitwiseZeroPayloadSHA256=digest(old_payload), LogsSHA256=log_hashes,
                  CornerRule='Exact concentric arcs; explicit straight connectors; no coordinate correction')
    if published:
        package = receipt('Package.json')
        equality = receipt('PublishedEquality.json')
        assert equality['Passed'] and not equality['Missing'] and not equality['Differences'] and not equality['UnexpectedCodeComponents']
        assert package['AcceptedCandidateSHA256'] == result['CandidateSHA256']
        book = ROOT / 'workbook/output/RC_Section_NDM.xlsm'
        assert digest(book.read_bytes()) == package['PublishedWorkbookSHA256']
        with zipfile.ZipFile(book) as archive:
            assert set(archive.namelist()) == set(package['UnchangedParts']) | {'xl/vbaProject.bin'}
            for name, expected in package['UnchangedParts'].items():
                assert digest(archive.read(name)) == expected, name
            assert digest(archive.read('xl/vbaProject.bin')) == package['VbaProjectSHA256']
        assert digest((ROOT / 'workbook/output/VBA_All_Code.txt').read_bytes()) == equality['ExportSHA256']
        for item in equality['Components']:
            paths = [path for folder in ('src', 'tests') for path in (ROOT / folder).rglob(item['Name'] + '.*')]
            assert len(paths) == 1 and digest(paths[0].read_bytes()) == item['SourceSHA256'], item['Name']
        published_text, published_hash, _ = inspect_log(FINAL / 'PublishedSmoke.txt', [EXPECTED[0], EXPECTED[1], EXPECTED[3]])
        assert published_hash == package['PublishedWorkbookSHA256']
        result.update(PublishedWorkbookSHA256=published_hash, SourceComponents=equality['SourceComponents'],
                      WorkbookComponents=equality['ExportComponents'], PreservedNonVBAParts=len(package['UnchangedParts']),
                      AllCodeSHA256=equality['ExportSHA256'], Published=True)
    else:
        result['Published'] = False
    (FINAL / 'Acceptance.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--published', action='store_true')
    print(json.dumps(verify(parser.parse_args().published), ensure_ascii=False))
