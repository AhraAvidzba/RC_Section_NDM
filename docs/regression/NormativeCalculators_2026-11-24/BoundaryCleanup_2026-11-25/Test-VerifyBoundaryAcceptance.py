"""ДЛЯ ТЕСТОВ: проверяет полноту адресной приёмки и точное совпадение расчётов."""
from pathlib import Path
import hashlib
import importlib.util
import json
import re
import subprocess
from zipfile import ZipFile

ROOT = next(p for p in Path(__file__).resolve().parents if (p / 'AGENTS.md').is_file())
D = Path(__file__).resolve().parent
read = lambda p: json.loads(p.read_text(encoding='utf-8-sig'))
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest().upper()

def module(name, filename):
    spec = importlib.util.spec_from_file_location(name, ROOT / 'tools/build_workbook' / filename)
    value = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(value)
    return value

package = read(D / 'PackageProof.json')
scope = read(D / 'SourceScope.json')
runs = read(D / 'DirectedRuns.json')
assert scope['Status'] == package['Status'] == 'PASS'
assert scope['UnchangedExistingMethodBodies'] == 333 and scope['NoNewVBAClassesOrModules']
assert package['NoNewVBAComponents'] and package['NativeComponents'] == 209
assert len(runs) == 6 and {(r['Label'], r['Mode']) for r in runs} == {
    (label, mode) for label in ('Baseline', 'Candidate', 'Boundary') for mode in ('No', 'Yes')}
assert sha(D / 'Baseline.xlsm') == package['BaselineSHA256']
assert sha(D / 'Candidate.xlsm') == package['CandidateSHA256']
assert sha(D / 'VBA_All_Code.txt') == package['ExportSHA256']
assert sha(ROOT / 'workbook/output/RC_Section_NDM.xlsm') == package['CandidateSHA256']
assert sha(ROOT / 'workbook/output/VBA_All_Code.txt') == package['ExportSHA256']
with ZipFile(D / 'Baseline.xlsm') as a, ZipFile(D / 'Candidate.xlsm') as b, ZipFile(D / 'CandidateBuild.xlsm') as compiled:
    assert a.namelist() == b.namelist()
    assert [n for n in a.namelist() if a.read(n) != b.read(n)] == ['xl/vbaProject.bin']
    assert b.read('xl/vbaProject.bin') == compiled.read('xl/vbaProject.bin')

preservation = module('boundary_source_review', 'Verify-NormativeCalculatorRefactor.py')
native = module('boundary_native_review', 'Package-NormativeCalculators.py')
original, current = native.components(D / 'OriginalNative.txt'), native.components(D / 'NativeCandidate.txt')
assert original.keys() == current.keys() and len(current) == 209
changed = {name for name in original if native.normalized(original[name][1]) != native.normalized(current[name][1])}
assert changed == set(package['ChangedComponents'])
unchanged = 0
changed_methods = []
for name in changed:
    paths = [p for folder in ('src', 'tests') for p in (ROOT / folder).rglob(name + '.*') if p.suffix in ('.cls', '.bas')]
    assert len(paths) == 1
    path = paths[0]
    source = path.read_text(encoding='utf-8-sig')
    body = re.search(r'(?ms)^Option Explicit.*', source)[0]
    if name == 'CStabilityCalculator':
        body = body.replace('1.797693134862315E+308', '1.79769313486231E+308')
    assert native.normalized(body) == native.normalized(current[name][1]), name
    old_source = subprocess.check_output(['git', 'show', '83f3803a:' + path.relative_to(ROOT).as_posix()], cwd=ROOT).decode('utf-8-sig')
    old_methods, new_methods = preservation.methods(old_source), preservation.methods(source)
    assert not old_methods.keys() - new_methods.keys()
    for key in old_methods:
        if preservation.canonical(old_methods[key]) == preservation.canonical(new_methods[key]):
            unchanged += 1
        else:
            changed_methods.append(path.relative_to(ROOT).as_posix() + ':' + key[0])
assert unchanged == 333 and sorted(changed_methods) == sorted(scope['ChangedExistingMethods'])
tracked = subprocess.check_output(['git', 'ls-tree', '-r', '--name-only', '83f3803a', 'src', 'tests'], cwd=ROOT).decode().splitlines()
assert {p for p in tracked if p.endswith(('.bas', '.cls'))} == {
    p.relative_to(ROOT).as_posix() for folder in ('src', 'tests') for p in (ROOT / folder).rglob('*') if p.suffix in ('.bas', '.cls')}

proofs = []
for run in runs:
    path = ROOT / run['Report']
    body = path.read_text(encoding='utf-8-sig')
    expected_hash = package['BaselineSHA256'] if run['Label'] == 'Baseline' else package['CandidateSHA256']
    assert run['Passed'] and run['SourceSHA256'] == expected_hash
    assert re.findall(r'^SUITE_STARTED: ([^;]+);', body, re.M) == run['Macros']
    assert re.findall(r'^SUITE_FINISHED: ([^;]+);', body, re.M) == run['Macros']
    assert len(run['Macros']) == (2 if run['Label'] == 'Boundary' else 8)
    assert not re.search(r'^(?:FAIL[: ]|RUNTIME ERROR|SCRIPT ERROR|WATCHDOG_FAILURE|RESTORE ERROR|CLEANUP ERROR)|failed=[1-9]', body, re.M)
    assert body.count('NORMATIVE_CORE_GUARD_RESTORED: True') == len(run['Macros'])
    assert 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' in body
    assert read(Path(str(path) + '.ProcessClosure.json'))['Passed']
    if run['Label'] == 'Boundary':
        assert re.search(r'^TOTAL_NORMATIVE_BOUNDARY: passed=13; failed=0\r?$', body, re.M)
        assert re.search(r'^TOTAL_NORMATIVE_PARTIAL_BRANCH: passed=11; failed=0\r?$', body, re.M)
    proofs.append({'Label': run['Label'], 'Mode': run['Mode'], 'Report': run['Report'],
                   'SHA256': sha(path), 'Executions': len(run['Macros']), 'Passed': True, 'ProcessClosed': True})

comparisons = []
for mode in ('No', 'Yes'):
    engineering = []
    for label in ('Baseline', 'Candidate'):
        run = next(r for r in runs if r['Label'] == label and r['Mode'] == mode)
        lines = (ROOT / run['Report']).read_text(encoding='utf-8-sig').splitlines()
        lines = [re.sub(r'; elapsedSec=[^;\r\n]+', '', x) for x in lines
                 if x.startswith(('OK:', 'TOTAL', 'PROJECTION_', 'STABILITY_CONFIG:'))]
        engineering.append(lines)
    assert engineering[0] == engineering[1], 'Engineering results differ: ' + mode
    comparisons.append({'Mode': mode, 'ExactlyMatchingEngineeringLines': len(engineering[0]), 'NumericToleranceUsedForComparison': False})

supplement = (ROOT / 'tools/build_workbook/Run-CurrentSupplementAcceptance.ps1').read_text(encoding='utf-8-sig')
macros = re.search(r'(?s)\$macros=@\((.*?)\n\)', supplement)[1]
assert len(re.findall(r"'modTest[^']+'", macros)) == 7
assert "'modTestCrackWidth.RunNormativeBoundaryTests'" in macros
assert supplement.count('ExpectedExecutions=14') == 2
coordinator = (ROOT / 'tools/build_workbook/Run-NormativeFullAcceptance.ps1').read_text(encoding='utf-8-sig')
assert '$extra.ExpectedExecutions -eq 14' in coordinator and 'RequiredSupplementExecutions=14' in coordinator
result = {'Status': 'DIRECTED_PASS', 'ComparisonExecutions': 32, 'BoundaryExecutions': 4,
          'Comparisons': comparisons, 'Reports': proofs, 'SourceProof': scope, 'PackageProof': package,
          'FutureFullSupplementIncludesBoundaryInBothModes': True, 'FullRunRepeated': False}
(D / 'Acceptance.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps({k: result[k] for k in ('Status', 'ComparisonExecutions', 'BoundaryExecutions', 'Comparisons', 'FullRunRepeated')}, ensure_ascii=False))
