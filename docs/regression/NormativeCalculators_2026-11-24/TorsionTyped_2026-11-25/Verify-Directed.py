"""ДЛЯ ТЕСТОВ: сверяет полный состав связанных проверок и прежние утверждения без нормализации чисел."""
import hashlib
import json
import re
from pathlib import Path

here = Path(__file__).resolve().parent
root = here.parents[3]


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def flatten(items):
    # Windows PowerShell сериализовал ранее накопленную пару как оболочку value/Count.
    result = []
    for item in items:
        result.extend(flatten(item['value']) if 'value' in item else [item])
    return result


runs = flatten(json.loads((here / 'DirectedRuns.json').read_text(encoding='utf-8-sig')))
assert len(runs) == 4
assert {(r['Label'], r['Mode']) for r in runs} == {
    (label, mode) for label in ('Baseline', 'Candidate') for mode in ('No', 'Yes')
}
proof = json.loads((here / 'PackageProof.json').read_text(encoding='utf-8-sig'))
details = []
reports = {}
for run in runs:
    label, mode = run['Label'], run['Mode']
    path = root / run['Report']
    body = path.read_text(encoding='utf-8-sig')
    assert run['Passed'] and len(run['Macros']) == 10
    assert re.findall(r'^SUITE_STARTED: ([^;]+);', body, re.M) == run['Macros']
    assert re.findall(r'^SUITE_FINISHED: ([^;]+);', body, re.M) == run['Macros']
    assert re.findall(r'^MODE_AFTER_SUITE: (.+)$', body, re.M) == [mode] * 10
    assert len(re.findall(r'^NORMATIVE_CORE_GUARD_RESTORED: True$', body, re.M)) == 10
    assert not re.search(r'^FAIL\b|Failed=[1-9]|failed=[1-9]|RUNTIME ERROR|WATCHDOG_FAILURE', body, re.M)
    assert 'SOURCE_UNCHANGED: True' in body
    assert 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' in body
    source_hash = digest(here / (label + '.xlsm'))
    assert source_hash == run['SourceSHA256'] == proof[label + 'SHA256']
    assert 'SOURCE_SHA256: ' + source_hash in body
    closure = json.loads(Path(str(path) + '.ProcessClosure.json').read_text(encoding='utf-8-sig'))
    assert closure['Passed'] and closure['ChildHasExited'] and closure['ChildExitCode'] == 0
    assert closure['Closure']['Passed'] and closure['RunnerCleanupValid']
    assert not closure['OrdinaryPathKilledAnything']
    reports[label, mode] = body
    totals = re.findall(r'^(\w+): Passed=(\d+); Failed=0$', body, re.M)
    assert len(totals) == 10
    details.append({'Label': label, 'Mode': mode, 'Macros': 10,
                    'Assertions': sum(int(n) for _, n in totals), 'Failures': 0,
                    'ReportSHA256': digest(path), 'SourceSHA256': source_hash})

for mode in ('No', 'Yes'):
    before, after = reports['Baseline', mode], reports['Candidate', mode]
    old_lines = re.findall(r'^PASS .+$', before, re.M)
    new_lines = re.findall(r'^PASS .+$', after, re.M)
    typed = [line for line in new_lines if line.startswith('PASS Typed ')]
    assert len(typed) == 65
    assert [line for line in new_lines if not line.startswith('PASS Typed ')] == old_lines
    old_totals = dict((s, int(n)) for s, n in re.findall(r'^(\w+): Passed=(\d+); Failed=0$', before, re.M))
    new_totals = dict((s, int(n)) for s, n in re.findall(r'^(\w+): Passed=(\d+); Failed=0$', after, re.M))
    assert old_totals.keys() == new_totals.keys()
    assert all(new_totals[s] == old_totals[s] + (65 if s == 'TorsionPost' else 0) for s in old_totals)

assert digest(root / 'workbook/output/RC_Section_NDM.xlsm') == proof['BaselineSHA256']
assert digest(here / 'VBA_All_Code.txt') == proof['ExportSHA256']
result = {'Status': 'PASS', 'Scope': '10 related torsion macros x Baseline/Candidate x No/Yes',
          'Executions': 40, 'NewFullRunsStarted': False, 'OriginalAssertionsIdentical': True,
          'NewTypedAssertionsPerMode': 65, 'NoNewVBAComponents': proof['NoNewVBAComponents'],
          'NativeCompilationComponents': proof['NativeComponents'],
          'AllNonVBAPartsByteIdentical': proof['AllNonVBAPartsByteIdentical'],
          'CandidateInstalled': False, 'Runs': details}
(here / 'DirectedAcceptance.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps(result, ensure_ascii=False))
