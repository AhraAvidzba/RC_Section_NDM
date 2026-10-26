"""ДЛЯ ТЕСТОВ: сводит только направленную приёмку; полный PASS не присваивает."""
from pathlib import Path
import hashlib, json, subprocess
from datetime import datetime, timezone

here = Path(__file__).resolve().parent
root = here.parents[2]
read = lambda name: json.loads((here / name).read_text(encoding='utf-8-sig'))
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest().upper()
book_sha = '44C63ECF5C69C783A9455ACA1314FF0FFBEC5F825D0585F58636967D155AD4E5'
code_sha = '252089A0A3CB25EB21A9C2A291646D33B1BAFB96A6CE08CD8EE1BB6A1A30C82D'
directed = read('FinalDirected04_Review.json')
assert directed['Passed'] and len(directed['Reports']) == 2
for report in directed['Reports']:
    assert report['FullPrimaryRead'] and report['SourceSHA256'] == book_sha
    assert len(report['Started']) == 12 and report['Started'] == report['Finished']
    assert sum(t['Passed'] for t in report['Totals']) == 4527
    assert not report['Errors'] and all(t['Failed'] == 0 for t in report['Totals'])
    body = Path(report['Path']).read_text(encoding='utf-8-sig')
    for case in ('expandedGeometryIs17', 'expandedGeometryRejected', 'expandedGeometryAddress',
                 'geometry16.restoredIxy', 'legacy15.localIxy', 'marker.mm', 'marker.cm', 'marker.m'):
        assert 'OK: postAudit03.contours.' + case in body
native = read('PublishedStructure01/Acceptance.json')
assert native['Passed'] and native['CheckCount'] == 85 and native['SourceSHA256'] == book_sha
assert not native['ImportPerformed'] and not native['MacrosRun'] and not native['SavePerformed']
equality = read('FinalCandidate07_SourceEquality.json')
assert equality['Passed'] and equality['SourceComponents'] == 175 and equality['ExportComponents'] == 181
preservation = read('ReleaseCandidate07/DataPreservation.json')
assert preservation['Status'] == 'PASS' and preservation['AfterSHA256'] == book_sha
layout = read('FreshBuild03/LayoutVerification.json')
assert layout['Status'] == 'PASS' and not layout['Failures']
workflow = read('Workflow03/Acceptance.json')
assert workflow['Passed'] and workflow['CompletedCases'] == 5
decimal = read('DecimalSnapshot02/PrimaryReview.json')
assert decimal['AssertionsPassed'] == 2392 and decimal['AssertionsFailed'] == 0 and not decimal['Issues']
cad = read('NativeCAD07/PrimaryReview.json')
assert cad['AssertionsPassed'] == 127 and cad['AssertionsFailed'] == 0 and not cad['Issues']
closure = read('ProcessClosureFinal.json')
assert closure['AllRecordedOwnedIdentitiesAbsent'] and not closure['Issues']
assert sha(root / 'workbook/output/RC_Section_NDM.xlsm') == book_sha
assert sha(root / 'workbook/output/VBA_All_Code.txt') == code_sha
assert sha(root / 'workbook/output/RC_Section_NDM_execution_report.txt') == '07DD18850E14E3306599D00DAA17A3D909B0FAFD530D149D4EDD62C972F3A340'
assert subprocess.check_output(['git','rev-parse','main'], cwd=root, text=True).strip() == '1a0c03a82ecb2fdbbabdbdb1e2bff8f554136516'
result = dict(Status='PASS_SCOPED_DIRECTED_CONTOUR_CELL_ACCEPTANCE', ObservedUTC=datetime.now(timezone.utc).isoformat(),
              WorkbookSHA256=book_sha, AllCodeSHA256=code_sha, DirectedNoYes=[4527,4527],
              SourceComponents=175, DocumentComponents=6, NativePublishedChecks=85,
              WorkflowCases=5, WorkflowAssertions=4298, NativeCADAssertions=127, DecimalAssertions=2392,
              FullBuildLayoutChecks=1219, UserDataPreserved=True, ForeignReportPreserved=True,
              OwnIdentitiesClosed=closure['UniqueOwnIdentityCount'], MainUnchanged=True,
              FullRegressionRun=False, FullRegressionPass=False,
              UnapprovedLongSuiteStarted=False, HyperlinkStrictDiagnostic='REVIEW_REQUIRED_164_XML_NORMALIZATION_DIFFERENCES',
              HyperlinkVisibleAppearanceVerified=True, Scope='Directed acceptance of feature branch; no 178+Focused full-suite claim')
(here/'Acceptance.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(result,ensure_ascii=False))
