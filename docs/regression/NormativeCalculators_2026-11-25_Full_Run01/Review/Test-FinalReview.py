"""ДЛЯ ТЕСТОВ: итоговая сверка завершённой приёмки и установленной книги."""
from datetime import datetime
import hashlib
import importlib.util
from io import BytesIO
import json
from pathlib import Path
import re
import subprocess
from zipfile import ZipFile

ROOT = next(p for p in Path(__file__).resolve().parents if (p / 'AGENTS.md').is_file())
FULL = ROOT / 'docs/regression/NormativeCalculators_2026-11-25_Full_Run01'
LOCAL = ROOT / 'docs/regression/NormativeCalculators_2026-11-24'
HASH = '646B9E6ACF78BA029136BED50561CB305D38B08982E6D00736DCA88B78611E57'
EXPORT_HASH = '05EB16BC8BE2D7908063ABF15704359F84CEEC9B5E2507A78E13D3BD2395D7D0'


def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def module(name, filename):
    spec = importlib.util.spec_from_file_location(name, ROOT / 'tools/build_workbook' / filename)
    value = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(value)
    return value


goal = read(FULL / 'GoalAcceptance.json')
assert goal['Status'] == 'PASS' and goal['SourceSHA256'] == HASH
assert goal['PrimaryExitCode'] == goal['SupplementExitCode'] == 0
run = read(FULL / 'Run.json')
assert run['SourceCommit'] == '721c0e9c52b641788ef910e254abd728a1b94cdd'
assert not run['SourceChanged'] and not run['RunnerChanged']
assert datetime.fromisoformat(goal['FinishedUTC']) < datetime.fromisoformat(goal['Deadline'])
core = module('core_final_review', 'Summarize-FullOnOffAcceptance.py').summarize(FULL)
assert core['status'] == 'PASS' and core['observedSuites'] == core['expectedSuites'] == 202
assert core['failedAssertions'] == 0 and not core['errors']
focused_run = read(FULL / 'Focused/Run.json')
focused = module('focused_final_review', 'Verify-FocusedAcceptance.py').verify(FULL / 'Focused', focused_run)
assert focused['Status'] == 'PASS' and focused['IndividualCases'] == 122 and not focused['Errors']
combined = module('combined_final_review', 'Finalize-CompleteAcceptance.py').finalize(FULL, FULL / 'Focused')
assert combined['Status'] == 'PASS' and combined['SameWorkbookBytes'] and combined['SingleInvocation']
independent = read(FULL / 'Focused/Independent.json')
assert independent['Passed'] and independent['IndependentIntegrals'] == 563
assert independent['ReviewedNumFailStates'] == len(independent['Failures']) == 7 and not independent['Unsupported']
assert all(x['Passed'] and x['Classification'] == 'ProvenImpossibleByNecessaryForceBounds' for x in independent['Failures'])
native = read(FULL / 'Focused/NativeCAD/Stage.json')
assert native['Passed'] and not native['UserDocumentsUsed'] and len(native['Results']) == 2

supplement = read(FULL / 'Supplement/Acceptance.json')
macros = [
    'modTestBatchCalculation.RunZeroLoadCombinationTests',
    'modTestTorsionPost.RunTorsionOnlyTests',
    'modTestWorkbookInterface.RunAudit03ExcelGuardTests',
    'modTestWorkbookInterface.RunTestCalculationFreshnessWorkflow',
    'modTestBatchCalculation.RunNormativeStabilityTests',
    'modTestBatchCalculation.RunNormativePartialBranchFailureTest',
]
totals = ['TOTAL_ZERO_LOAD_COMBINATIONS', 'TorsionOnly', 'TOTAL_AUDIT03_EXCEL_GUARD',
          'TOTAL_FRESHNESS_WORKFLOW', 'TOTAL_NORMATIVE_STABILITY', 'TOTAL_NORMATIVE_PARTIAL_BRANCH']
assert supplement['Status'] == 'PASS' and supplement['ExpectedExecutions'] == 12
assert supplement['SourceSHA256'] == HASH
assert [x['Mode'] for x in supplement['Results']] == ['No', 'Yes']
supplement_proofs = []
for item in supplement['Results']:
    path = ROOT / item['Report']
    assert sha(path) == item['ReportSHA256'] and item['Passed'] and item['ExitCode'] == 0
    assert item['Macros'] == macros and not item['Errors'] and item['ProcessClosed']
    body = path.read_text(encoding='utf-8-sig')
    assert re.findall(r'^SUITE_STARTED: ([^;]+);', body, re.M) == macros
    assert re.findall(r'^SUITE_FINISHED: ([^;]+);', body, re.M) == macros
    assert not re.search(r'^(?:FAIL[: ]|RUNTIME ERROR|SCRIPT ERROR|WATCHDOG_FAILURE|RESTORE ERROR|CLEANUP ERROR)', body, re.M)
    lines = []
    for total in totals:
        matches = re.findall(r'^' + total + r': [^\r\n]+', body, re.M)
        assert len(matches) == 1
        assert re.search(r'(?:failed|Failed)=0(?:;|$)', matches[0])
        lines += matches
    assert 'TOTAL_NORMATIVE_PARTIAL_BRANCH: passed=11; failed=0' in lines
    assert read(Path(str(path) + '.ProcessClosure.json'))['Passed']
    supplement_proofs.append({'Mode': item['Mode'], 'Totals': lines, 'Closed': True})

book = ROOT / 'workbook/output/RC_Section_NDM.xlsm'
export = ROOT / 'workbook/output/VBA_All_Code.txt'
assert sha(book) == sha(FULL / 'Source/RC_Section_NDM.xlsm') == sha(LOCAL / 'Accepted.xlsm') == HASH
assert sha(export) == sha(LOCAL / 'VBA_All_Code.txt') == EXPORT_HASH
original = subprocess.check_output(['git', 'show', '6e8442d2:workbook/output/RC_Section_NDM.xlsm'], cwd=ROOT)
assert hashlib.sha256(original).hexdigest().upper() == '2D413237AAF72D2D134290268F828DBC99140C2850303B558D4B1F68C96AAE06'
with ZipFile(BytesIO(original)) as previous, ZipFile(book) as current, ZipFile(LOCAL / 'FinalCandidate.xlsm') as compiled:
    assert previous.namelist() == current.namelist()
    changed = [n for n in previous.namelist() if previous.read(n) != current.read(n)]
    assert changed == ['xl/vbaProject.bin']
    assert current.read('xl/vbaProject.bin') == compiled.read('xl/vbaProject.bin')
sources = module('source_final_review', 'Package-NormativeCalculators.py').source_equality(LOCAL)
assert sources['Passed'] and sources['NativeComponents'] == 209
proof = read(FULL / 'Review/ExpressionPreservation.json')
assert proof['Status'] == 'PASS' and proof['Methods'] == 239 and proof['ExistingBatchTestMethodsUnchanged'] == 274
classes = {}
for filename in ['Crack/CCrackWidthSP35Calculator.cls', 'Crack/CCrackWidthSP63Calculator.cls',
                 'Stability/CStabilitySP35Calculator.cls', 'Stability/CStabilitySP63Calculator.cls']:
    text = (ROOT / 'src' / filename).read_text(encoding='utf-8-sig')
    assert 'Excel.' not in text and 'Worksheet' not in text
    classes[Path(filename).stem] = len(text.splitlines())
assert min(classes.values()) >= 110
local = read(LOCAL / 'AcceptanceAndPackage.json')
assert local['Acceptance']['Status'] == 'DIRECTED_PASS' and local['Acceptance']['Executions'] == 76

requirements = [
    {'Requirement': 'SP35/SP63 cracks and stability separated into substantive classes', 'Evidence': classes, 'Passed': True},
    {'Requirement': 'Calculation expressions, public preparation and existing test expectations preserved', 'Evidence': {'Methods': 239, 'ExistingBatchTestMethods': 274, 'SourceEquality': sources}, 'Passed': True},
    {'Requirement': 'Local baseline/candidate acceptance in No/Yes', 'Evidence': 'docs/regression/NormativeCalculators_2026-11-24/AcceptanceAndPackage.json and LocalEvidence.zip', 'Passed': True},
    {'Requirement': 'Final workbook and export installed; formatting, data, formulas and lists preserved', 'Evidence': {'BookSHA256': HASH, 'ExportSHA256': EXPORT_HASH, 'ChangedZIPParts': changed, 'NativeVBAEqualsCompiledCandidate': True}, 'Passed': True},
    {'Requirement': 'All primary No/Yes groups', 'Evidence': {'Groups': len(run['CompletedJobs']), 'Executions': core['observedSuites'], 'FailedAssertions': core['failedAssertions']}, 'Passed': True},
    {'Requirement': 'Mandatory focused plan, real CAD, independent control and save/reopen', 'Evidence': {'SingleCases': focused['IndividualCases'], 'Batch15BothModes': True, 'StabilityBothModes': True, 'AsymmetricSteelBothModes': True, 'NativeCAD': native, 'IndependentIntegrals': 563, 'ReviewedNumFailStates': 7}, 'Passed': True},
    {'Requirement': 'All 12 supplements, including four newly normative executions', 'Evidence': supplement_proofs, 'Passed': True},
    {'Requirement': 'One complete invocation on unchanged bytes, finished before agreed deadline', 'Evidence': {'SourceSHA256': HASH, 'StartedUTC': run['StartedUTC'], 'FinishedUTC': goal['FinishedUTC'], 'Deadline': goal['Deadline'], 'RepeatRequired': False}, 'Passed': True},
]
result = {'Status': 'PASS', 'ReviewedUTC': datetime.now().astimezone().isoformat(),
          'SourceCommit': run['SourceCommit'], 'Requirements': requirements,
          'RemainingDeliveryActions': ['Commit final evidence and report', 'Push codex/normative-calculators', 'Pause existing ndm automation', 'Mark active goal complete']}
(FULL / 'Review/FinalReview.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'Status': 'PASS', 'Requirements': len(requirements), 'Core': 202, 'Focused': 122, 'Supplement': 12, 'Classes': classes}, ensure_ascii=False))
