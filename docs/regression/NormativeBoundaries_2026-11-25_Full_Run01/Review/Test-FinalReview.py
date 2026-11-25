"""ДЛЯ ТЕСТОВ: сверка первичных журналов полной и последующей адресной приёмки."""
import hashlib
import importlib.util
import json
from pathlib import Path
import re

ROOT = next(p for p in Path(__file__).resolve().parents if (p / 'AGENTS.md').is_file())
FULL = Path(__file__).resolve().parent.parent
TYPED = ROOT / 'docs/regression/NormativeCalculators_2026-11-24/TorsionTyped_2026-11-25'
BASE = 'F06EE141AFE777FCE1353099AA575E1D0D4BCFCF4E6E111929BE175D70D9AB8B'


def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def module(name, file):
    spec = importlib.util.spec_from_file_location(name, ROOT / 'tools/build_workbook' / file)
    value = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(value)
    return value


goal = read(FULL / 'GoalAcceptance.json')
assert goal['Status'] == 'PASS' and goal['SourceSHA256'] == BASE
assert goal['PrimaryExitCode'] == goal['SupplementExitCode'] == 0
core = module('primary_review', 'Summarize-FullOnOffAcceptance.py').summarize(FULL)
assert core['status'] == 'PASS' and not core['errors']
assert core['observedSuites'] == core['completedSuites'] == core['expectedSuites'] == 202
assert core['failedAssertions'] == 0
focused_run = read(FULL / 'Focused/Run.json')
focused = module('focused_review', 'Verify-FocusedAcceptance.py').verify(FULL / 'Focused', focused_run)
assert focused['Status'] == 'PASS' and focused['IndividualCases'] == 122 and not focused['Errors']
combined = module('combined_review', 'Finalize-CompleteAcceptance.py').finalize(FULL, FULL / 'Focused')
assert combined['Status'] == 'PASS' and combined['SingleInvocation'] and combined['SameWorkbookBytes']
native = read(FULL / 'Focused/NativeCAD/Stage.json')
assert native['Passed'] and not native['UserDocumentsUsed'] and len(native['Results']) == 2
independent = read(FULL / 'Focused/Independent.json')
assert independent['Passed'] and independent['IndependentIntegrals'] == 563
assert independent['ReviewedNumFailStates'] == len(independent['Failures']) == 7
assert not independent['Unsupported']
assert all(x['Passed'] and x['Classification'] == 'ProvenImpossibleByNecessaryForceBounds'
           for x in independent['Failures'])

supplement = read(FULL / 'Supplement/Acceptance.json')
assert supplement['Status'] == 'PASS' and supplement['ExpectedExecutions'] == 14
assert supplement['SourceSHA256'] == BASE
assert [x['Mode'] for x in supplement['Results']] == ['No', 'Yes']
macros = [
    'modTestBatchCalculation.RunZeroLoadCombinationTests',
    'modTestTorsionPost.RunTorsionOnlyTests',
    'modTestWorkbookInterface.RunAudit03ExcelGuardTests',
    'modTestWorkbookInterface.RunTestCalculationFreshnessWorkflow',
    'modTestBatchCalculation.RunNormativeStabilityTests',
    'modTestBatchCalculation.RunNormativePartialBranchFailureTest',
    'modTestCrackWidth.RunNormativeBoundaryTests',
]
for item in supplement['Results']:
    path = ROOT / item['Report']
    assert sha(path) == item['ReportSHA256']
    assert item['Passed'] and item['ExitCode'] == 0 and item['ProcessClosed'] and not item['Errors']
    assert item['Macros'] == macros
    body = path.read_text(encoding='utf-8-sig')
    assert re.findall(r'^SUITE_STARTED: ([^;]+);', body, re.M) == macros
    assert re.findall(r'^SUITE_FINISHED: ([^;]+);', body, re.M) == macros
    assert not re.search(r'^FAIL\b|Failed=[1-9]|failed=[1-9]|RUNTIME ERROR|WATCHDOG_FAILURE|RESTORE ERROR|CLEANUP ERROR', body, re.M)
    assert 'SOURCE_UNCHANGED: True' in body
    assert read(Path(str(path) + '.ProcessClosure.json'))['Passed']
for folder in (FULL, FULL / 'Focused'):
    for path in folder.glob('*.txt.ProcessClosure.json'):
        closure = read(path)
        assert closure['Passed'] and closure['ChildHasExited'] and closure['ChildExitCode'] == 0

typed = read(TYPED / 'DirectedAcceptance.json')
package = read(TYPED / 'PackageProof.json')
assert typed['Status'] == 'PASS' and typed['Executions'] == 40
assert typed['OriginalAssertionsIdentical'] and typed['NewTypedAssertionsPerMode'] == 65
assert not typed['NewFullRunsStarted'] and typed['NoNewVBAComponents']
assert sha(TYPED / 'Candidate.xlsm') == package['CandidateSHA256']
assert sha(TYPED / 'VBA_All_Code.txt') == package['ExportSHA256']
assert sha(ROOT / 'workbook/output/RC_Section_NDM.xlsm') == BASE
assert sha(ROOT / 'workbook/output/VBA_All_Code.txt') == 'E9998E65CFA49758C0A0334560A3760A1190C787C28113D74C357814B12D623E'
result = {
    'Status': 'PASS', 'FullSourceCommit': core['sourceCommit'], 'FullWorkbookSHA256': BASE,
    'FullExecutions': 202, 'FullPassedAssertions': combined['CorePassedAssertions'],
    'FocusedSingleCases': 122, 'FocusedPassedAssertions': focused['PassedAssertions'],
    'FocusedSeconds': focused['ElapsedSeconds'], 'RealAutoCADStages': 2,
    'IndependentIntegrals': 563, 'ReviewedNumFailStates': 7, 'SupplementExecutions': 14,
    'SubsequentTypedDirectedExecutions': 40, 'TypedCandidateSHA256': package['CandidateSHA256'],
    'NoNewFullRunForTypedChange': True, 'ReadyForInstallation': True,
    'Scope': 'Full PASS belongs to F06 baseline; subsequent typed torsion candidate accepted only by directed tests as authorized.'
}
(FULL / 'Review/FinalReview.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps(result, ensure_ascii=False))
