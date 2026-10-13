"""Bind every named ValidationPlan VBA scenario to current ON/OFF evidence."""
import argparse
import importlib.util
import json
import re
from pathlib import Path

spec = importlib.util.spec_from_file_location('logs', Path(__file__).with_name('Summarize-GeometryFull.py'))
logs = importlib.util.module_from_spec(spec)
spec.loader.exec_module(logs)

DIRECT = {
    'RunSectionSolverTests': 'Full', 'RunCapacitySolverTests': 'Full',
    'RunGeometryQueryTests': 'Supplement', 'RunContourMeshTests': 'Domain',
    **dict.fromkeys(('RunPlateauSlopeTests', 'RunAudit03MaterialControlUnitsTests',
                    'RunPlateauResultsDiagramTests', 'RunConcreteMeshCenterTests',
                    'RunCommonStrainLimitTests', 'RunAsymmetricSteelStrengthTests',
                    'RunTangentClipRegressionTests', 'RunSolverWorkStatsTests',
                    'RunBatchWorkStatsTests'), 'Extras'),
    **dict.fromkeys(('RunAudit03UltimateContracts', 'RunCapacityMomentMeritTests',
                    'RunLimitGuessScalingTests', 'RunBracketFinalizationRecoveryTests',
                    'RunSavedCrackLambdaToleranceTests', 'RunAudit03LambdaSettingsTests',
                    'RunExecutionReportFormattingTests'), 'ValidationPlan'),
}
ALIASES = {
    'RunContinuationStateTests': 'RunSectionSolverTests',
    'RunUltimateFrameRegressionTests': 'RunCapacitySolverTests',
    'RunNumericalPathPreparationTests': 'RunCapacitySolverTests',
    'RunProjectedLimitRecoveryTests': 'RunCapacitySolverTests',
    'RunUltimateStepScaleRegressionTests': 'RunCapacitySolverTests',
    'RunAudit03UltimateStagnation': 'RunAudit03UltimateContracts',
}
NATIVE = ('RunRealAutoCADContourMeshTests', 'RunRealAutoCADFarOriginTests', 'RunRealAutoCADFarCapacityTests')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory', type=Path, required=True)
    args = parser.parse_args()
    root = args.directory
    repo = Path(__file__).resolve().parents[2]
    plan = repo/'docs/ValidationPlan.md'
    named = list(dict.fromkeys(re.findall(r'\bRun[A-Z]\w+\b', plan.read_text(encoding='utf-8-sig'))))
    errors, methods, rows = [], {}, []
    for path in (repo/'tests').glob('*.bas'):
        text = path.read_text(encoding='utf-8-sig')
        for match in re.finditer(r'(?ms)^(?:Public|Private) (Function|Sub) (\w+)\b.*?^End \1\b', text):
            calls = set()
            for line in match.group().splitlines():
                line = re.sub(r'"(?:[^"]|"")*"', '""', line).split("'", 1)[0]
                calls.update(re.findall(r'(?:^|:)\s*(Test\w+)\s+', line))
            methods[(path.stem, match[2])] = {'body': match.group(), 'calls': calls, 'source': path}

    def unique(name):
        values = [(key, value) for key, value in methods.items() if key[1] == name]
        if len(values) != 1:
            errors.append('Expected one test entrypoint: ' + name)
            return None, None
        return values[0]

    def reachable(module, name, seen=None):
        seen = set() if seen is None else seen
        method = methods.get((module, name))
        for call in method['calls'] if method else ():
            if call not in seen:
                seen.add(call)
                reachable(module, call, seen)
        return seen

    expected_counts = {'Full': 8, 'Supplement': 14, 'Extras': 9, 'Domain': 21, 'ValidationPlan': 7}
    reports = {(group, mode): logs.inspect(root/f'Final{group}_{mode}.txt', logs.sha(root/('ReleaseLambdaFixed.xlsm' if mode == 'On' and group in ('Supplement', 'Extras') else 'FinalReady.xlsm')), count)
               for group, count in expected_counts.items() for mode in ('Off', 'On')}
    for mode in ('Off', 'On'):
        scope = logs.full_scope(root, mode)
        reports[('Full', mode)] = {**scope, 'suites': scope['acceptedSuites']}
    native = json.loads((root/'NativeFinalEvidence.json').read_text(encoding='utf-8-sig'))
    for name in named:
        row = {'entrypoint': name}
        if name in NATIVE:
            actual = [r for r in native.get('reports', []) if r['macro'] == name]
            row['evidence'] = [r['path'] for r in actual]
            row['status'] = 'PASS' if len(actual) == 2 and {r['mode'] for r in actual} == {'Off', 'On'} and all(r['status'] == 'PASS' for r in actual) else 'INCOMPLETE_OR_FAILED'
        else:
            target = ALIASES.get(name, name)
            group = DIRECT.get(target)
            key, method = unique(name)
            target_key, target_method = unique(target)
            subset = True
            if name in ALIASES:
                reached = reachable(target_key[0], target) if target_key else set()
                subset = bool(method) and bool(method['calls']) and method['calls'] <= reached
                row.update(coveredBy=target, privateCases=sorted(method['calls']) if method else [], privateCasesReachable=subset)
            if group is None:
                errors.append('Unmapped ValidationPlan entrypoint: ' + name)
                row['status'] = 'UNMAPPED'
            else:
                proofs = [reports[(group, mode)] for mode in ('Off', 'On')]
                actual_target = '.'.join(target_key) if target_key else ''
                executed = all(r['status'] == 'PASS' and any(s['name'] == actual_target and s['status'] == 'PASS' for s in r['suites']) for r in proofs)
                evidence = sorted({s.get('report', r['name']) for r in proofs for s in r['suites'] if s['name'] == actual_target})
                row.update(evidence=evidence, status='PASS' if subset and executed else 'INCOMPLETE_OR_FAILED')
            if method:
                row['source'] = method['source'].relative_to(repo).as_posix()
                row['sourceSHA256'] = logs.sha(method['source'])
        rows.append(row)
    if any(row['status'] != 'PASS' for row in rows):
        errors.append('Not all named scenarios are proven in both modes')
    result = {'status': 'PASS' if not errors else 'INCOMPLETE_OR_FAILED', 'validationPlanSHA256': logs.sha(plan),
              'scenarios': rows, 'errors': errors, 'aliasPolicy': 'Same private test bodies reached from a completed current aggregator; no substitution with similar tests'}
    (root/'ValidationPlanCoverage.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps({'status': result['status'], 'scenarios': len(rows), 'errors': errors}))
    return bool(errors)


if __name__ == '__main__':
    raise SystemExit(main())
