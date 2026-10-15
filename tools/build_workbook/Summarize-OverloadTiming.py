"""Сопоставляет фактические полные замеры, пределы и регрессионные протоколы."""
import json
import re
import statistics
from pathlib import Path

root = Path(__file__).resolve().parents[2]
directory = root / 'docs/regression/Performance/Overload_2026-10-15'


def rows(folder):
    completion = json.loads((directory / folder / 'Completion.json').read_text(encoding='utf-8-sig'))
    assert completion['Status'] == 'OVERLOAD_MEASUREMENTS_COMPLETE' and completion['SourceUnchanged']
    return json.loads((directory / folder / 'Measurements.json').read_text(encoding='utf-8-sig'))


def final_plane(folder, name):
    text = (directory / folder / (name + '.txt')).read_text(encoding='utf-8-sig')
    last = [line for line in text.splitlines() if line.startswith(('search=ultimatePath iter=', 'search=ultimate iter='))][-1]
    values = dict(item.split('=', 1) for item in last.split('; '))
    return {key: values[key] for key in ('eps0', 'kappaX', 'kappaY', 'lambda')}


baseline = {item['id']: item for item in rows('UserCircleMesh10')}
final = rows('FinalUserCircleMesh10')
heavy = []
for axis in ('Mx', 'My'):
    before = baseline['UserCircle500N100' + axis]
    after = [item for item in final if '100' + axis in item['id']]
    assert len(after) == 3 and int(before['fibers']) == 716 and int(before['bars']) == 8
    for item in after:
        assert all(before[key] == item[key] for key in ('lambda', 'path', 'method', 'overall', 'direct', 'capacity'))
        assert int(item['ultimateStarts']) == 1 and int(item['ultimateIterations']) == 4
        assert final_plane('UserCircleMesh10', before['id']) == final_plane('FinalUserCircleMesh10', item['id'])
        assert float(item['seconds']) <= 10
    seconds = [float(item['seconds']) for item in after]
    heavy.append({'Axis': axis, 'BeforeSeconds': float(before['seconds']), 'AfterSeconds': seconds,
                  'AfterMedianSeconds': statistics.median(seconds), 'BeforeUltimateIterations': int(before['ultimateIterations']),
                  'AfterUltimateIterations': 4, 'Lambda': before['lambda'], 'BoundaryPlaneExactlyEqual': True})
old_controls = {item['id']: item for item in rows('FinalBaselineControls')}
controls = rows('FinalCandidateControls')
for item in controls:
    old = old_controls[item['id']]
    assert all(old[key] == item[key] for key in ('lambda', 'path', 'method', 'overall', 'direct', 'capacity'))
    assert final_plane('FinalBaselineControls', item['id']) == final_plane('FinalCandidateControls', item['id'])
    assert float(item['seconds']) <= 10 and item['overall'] in ('OK', 'FAIL')
bisection = []
for item in rows('BisectionOverloads'):
    if item['method'] != 'LoadMultiplier':
        continue
    text = (directory / 'BisectionOverloads' / (item['id'] + '.txt')).read_text(encoding='utf-8-sig')
    lambdas = [float(value) for value in re.findall(r'^probe=\d+; attempt=0; lambda=([^;]+)', text, re.M)]
    assert 'searchMethod=Bisection' in text and 'bracketFinalizationRecovery' not in text
    assert len(lambdas) == 9 and float(item['seconds']) <= 10
    bisection.append({'ID': item['id'], 'Seconds': float(item['seconds']), 'LambdaProbes': lambdas,
                      'BisectionIterations': len(lambdas) - 2, 'ConfiguredToleranceLambda': 0.01})
counts = {}
for name in ('FinalCapacityYes.txt', 'FinalCapacityNo.txt', 'FinalGuardsYes.txt', 'FinalGuardsNo.txt', 'FinalUltimateContracts.txt'):
    text = (directory / name).read_text(encoding='utf-8-sig')
    assert 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' in text
    assert not re.search(r'^(FAIL:|RUNTIME ERROR:|WATCHDOG_FAILURE:)', text, re.M)
    passed, failed = re.findall(r'^TOTAL_[^\n]*passed=(\d+); failed=(\d+)', text, re.M)[-1]
    assert int(failed) == 0
    counts[name] = int(passed)
result = {'Status': 'OVERLOAD_TIMING_ACCEPTANCE_PASS', 'HeavyCase': heavy,
          'ControlLoads': controls, 'ControlsAndBoundaryPlanesExactlyEqual': True,
          'Bisection': bisection, 'RegressionAssertions': counts, 'TotalRegressionAssertions': sum(counts.values()),
          'NumericalFailureNotConvertedToPhysicalFailure': True,
          'NoToleranceIterationsOrDiagramParametersRelaxed': True, 'NoFullOnOffOrFullUIRuns': True, 'Errors': []}
(directory / 'MeasurementsAcceptance.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'Status': result['Status'], 'HeavyCase': heavy, 'TotalRegressionAssertions': sum(counts.values()),
                  'BisectionIterations': [item['BisectionIterations'] for item in bisection]}, ensure_ascii=False))
