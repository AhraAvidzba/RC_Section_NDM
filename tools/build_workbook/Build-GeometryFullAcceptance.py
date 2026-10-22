"""Aggregate completed geometry acceptance; missing evidence is never PASS."""
import argparse
import hashlib
import json
import re
import subprocess
import sys
from pathlib import Path

import importlib.util

spec = importlib.util.spec_from_file_location('full_logs', Path(__file__).with_name('Summarize-GeometryFull.py'))
logs = importlib.util.module_from_spec(spec)
spec.loader.exec_module(logs)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory', type=Path, required=True)
    args = parser.parse_args()
    root = args.directory
    errors, artifacts, groups = [], {}, []

    def artifact(relative):
        path = root / relative
        if not path.is_file():
            errors.append('Missing: ' + relative)
            return None
        artifacts[relative] = {'path': relative, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()}
        return path

    def read(relative):
        path = artifact(relative)
        return json.loads(path.read_text(encoding='utf-8-sig')) if path else {}

    def group(name, passed, details):
        groups.append({'name': name, 'status': 'PASS' if passed else 'INCOMPLETE_OR_FAILED',
                       'modes': ['Off', 'On'], 'details': details})
        if not passed:
            errors.append('Group incomplete: ' + name)

    candidate = read('FinalReleaseCandidate.json')
    if candidate.get('book') != 'ReleaseSnapshotFixed.xlsm' or candidate.get('dump') != 'VBA_SnapshotFixed.txt':
        errors.append('Unexpected release candidate')
    release, tested = root/'ReleaseSnapshotFixed.xlsm', root/'FinalReady.xlsm'
    release_hash, tested_hash = logs.sha(release), logs.sha(tested)
    lambda_hash = logs.sha(root/'ReleaseLambdaFixed.xlsm')
    for proof in ('AuxiliaryInputsEvidence.json', 'ModeAuxiliaryInputsEvidence.json'):
        auxiliary = read(proof)
        if auxiliary.get('status') != 'PASS' or len(auxiliary.get('inputs', [])) != 2:
            errors.append('Private auxiliary source equality is missing: ' + proof)
        helper = Path(auxiliary.get('helperSource', ''))
        if not helper.is_file() or logs.sha(helper) != auxiliary.get('helperSourceSHA256'):
            errors.append('Auxiliary helper source changed: ' + proof)
        for value in auxiliary.get('inputs', []):
            for path_key, hash_key in (('book', 'bookSHA256'), ('dump', 'dumpSHA256'), ('referenceDump', 'referenceDumpSHA256')):
                path = artifact(value[path_key])
                if not path or logs.sha(path) != value[hash_key]:
                    errors.append('Auxiliary proof input changed: ' + value[path_key])
            if value.get('components') != 150 or value.get('changed') or value.get('added') != ['modTestFinalGeometryChecks'] or not value.get('helperEqual'):
                errors.append('Auxiliary contains unverified code changes')
    full = read('FinalFullEvidence.json')
    for name, prefix, count in (('Full8', 'Full', 8), ('Geometry14', 'Supplement', 14), ('ValidationExtras9', 'Extras', 9)):
        reports = ([logs.full_scope(root, mode) for mode in ('Off', 'On')] if prefix == 'Full' else
                   [logs.inspect(root/f'Final{prefix}_{mode}.txt', tested_hash if mode == 'Off' else lambda_hash, count) for mode in ('Off', 'On')])
        for r in reports:
            if prefix != 'Full':
                artifact(r['name'])
        group(name, all(r['status'] == 'PASS' for r in reports), reports)
    for name in ('FinalFull_Off.txt', 'FinalFull_On.txt', 'LambdaFixedBatch_On.txt',
                 'LambdaFixedDirectedRetry_Off.txt', 'LambdaFixedDirectedRetry_On.txt', 'LambdaBaselineProbe_On.txt',
                 'ProfileSnapshotFixed_Off.txt', 'ProfileSnapshotFixed_On.txt', 'ProfileBaselineOriginal_On.txt'):
        artifact(name)
    if full.get('status') != 'PASS' or full.get('sourceSHA256') != tested_hash:
        errors.append('Full aggregate is not the accepted tested input')
    validation_plan = [logs.inspect(root/f'FinalValidationPlan_{mode}.txt', tested_hash, 7) for mode in ('Off', 'On')]
    for value in validation_plan:
        artifact(value['name'])
    group('ValidationPlan7', all(r['status'] == 'PASS' for r in validation_plan), validation_plan)
    plan_coverage = read('ValidationPlanCoverage.json')
    plan_path = Path(__file__).resolve().parents[2]/'docs/ValidationPlan.md'
    if plan_coverage.get('status') != 'PASS' or plan_coverage.get('validationPlanSHA256') != logs.sha(plan_path) or len(plan_coverage.get('scenarios', [])) != 29:
        errors.append('Complete current ValidationPlan coverage is unproven')

    domain = read('FinalDomainEvidence.json')
    for r in domain.get('reports', []):
        artifact(r['name'])
    group('Domain21', domain.get('status') == 'PASS' and len(domain.get('reports', [])) == 2, domain)

    family_checks = []
    for mode in ('Off', 'On'):
        for code in ('SP35', 'SP63'):
            name = f'FinalFamily_{code}_{mode}_Equality.json'
            baseline, actual = root/f'FinalFamily_{code}_{mode}_Baseline.txt', root/f'FinalFamily_{code}_{mode}_Current.txt'
            if baseline.is_file() and actual.is_file():
                comparison = subprocess.run([sys.executable, str(Path(__file__).with_name('Compare-GeometrySnapshots.py')),
                                             str(baseline), str(actual), '--exact', '--report', str(root/name)], capture_output=True, text=True)
                if comparison.returncode:
                    errors.append('Fresh family comparison failed: ' + code + '/' + mode)
            value = read(name)
            family_checks.append(value.get('status') == 'PASS' and value.get('exact') and value.get('records') == 48114 and len(value.get('fixtures', [])) == 9)
            for kind in ('Baseline', 'Current'):
                artifact(f'FinalFamily_{code}_{mode}_{kind}.txt')
            artifact(f'FinalFamilyComplement_{code}_{mode}.txt')
            complement = logs.inspect(root/f'FinalFamilyComplement_{code}_{mode}.txt', logs.sha(root/'AuxiliaryFinal.xlsm'), 1)
            if complement['status'] != 'PASS':
                errors.append('Family complement incomplete: ' + code + '/' + mode)
        artifact(f'FinalSP35Projection_{mode}.txt')
        projection_path = root/f'FinalSP35Projection_{mode}.txt'
        projection = logs.inspect(projection_path, logs.sha(root/'AuxiliaryModeFinal.xlsm'), 1) if (root/'AuxiliaryModeFinal.xlsm').is_file() else {'status': 'NOT_RUN'}
        if projection['status'] != 'PASS':
            errors.append('SP35 projection incomplete: ' + mode)
        if not projection_path.is_file() or ('PROJECTION_DIAGRAM_MODE: ' + ('No' if mode == 'Off' else 'Yes')) not in projection_path.read_text(encoding='utf-8-sig'):
            errors.append('SP35 projection actual mode is not proven: ' + mode)
    group('Family9_SP35_SP63', all(family_checks) and len(family_checks) == 4, 'Four exact comparisons, nine families per norm/mode')

    native = read('NativeFinalEvidence.json')
    for r in native.get('reports', []):
        path = artifact(r['path'])
        if not path or logs.sha(path) != r.get('sha256') or r.get('status') != 'PASS':
            errors.append('Native report changed: ' + r['path'])
        directory = str(Path(r['path']).parent).replace('\\', '/')
        for suffix in ('NativeInvocation.json', 'NativeCADOwnership.json', 'NativeCADIdentity.json', 'NativeExcelOwnership.json'):
            artifact(directory + '/' + suffix)
        invocation = read(directory + '/NativeInvocation.json')
        ownership = read(directory + '/NativeCADOwnership.json')
        if not invocation.get('Passed') or not invocation.get('SourceUnchanged') or ownership.get('UserDocumentsUsed') is not False:
            errors.append('Native ownership/source guard failed: ' + directory)
    group('NativeCAD5', native.get('status') == 'PASS' and len(native.get('reports', [])) == 10, native)

    one_class = []
    one_source = logs.sha(root/'OneClassFinal.xlsm')
    for mode in ('Off', 'On'):
        names = [(f'OneClassFinalInputBuild_{mode}.txt', 3)] + [(f'OneClassFinalWorkflow_{mode}_{width}.txt', 1) for width in ('1200', '1400')]
        for name, count in names:
            artifact(name)
            result = logs.inspect(root/name, one_source, count)
            one_class.append(result)
            if 'Workflow' in name and (root/name).is_file():
                text = (root/name).read_text(encoding='utf-8-sig')
                if not re.search(r'(?m)^RESULTS_SAVE_REOPEN: .*equal=True', text) or not re.search(r'(?m)^STATUS_STYLE_SAVE_REOPEN: .*equal=True', text):
                    errors.append('One-class reopen proof missing: ' + name)
    group('OneClass', all(r['status'] == 'PASS' for r in one_class), one_class)

    performance = read('FinalPerformanceEvidence.json')
    visual = read('FinalVisualEvidence.json')
    group('WorkflowPerformance', performance.get('status') == 'PASS' and len(performance.get('comparisons', [])) == 18, performance)
    group('PlotModes', visual.get('status') == 'PASS' and len(visual.get('images', [])) == 88
          and len(visual.get('pairs', [])) == 36 and visual.get('visualReview', {}).get('status') == 'REVIEWED', visual)
    for value in performance.get('artifacts', []) + visual.get('images', []) + visual.get('contactSheets', []):
        path = artifact(value['path'])
        if not path or logs.sha(path) != value['sha256']:
            errors.append('Performance or inspected image changed: ' + value['path'])
    if visual.get('status') == 'PASS':
        artifact('VisualQA/Review.json')
    for mode in ('Off', 'On'):
        for kind in ('Baseline', 'Current'):
            for prefix in ('FinalWorkflow', 'FinalStagesMixed'):
                artifact(f'{prefix}_{mode}_{kind}.txt')
            directory = root/f'FinalPlots_{mode}_{kind}'
            for file in directory.glob('*.png'):
                artifact(str(file.relative_to(root)).replace('\\', '/'))
        artifact(f'FinalPlotModes_{mode}.txt')
        for file in (root/f'FinalPlotModes_{mode}').glob('*.png'):
            artifact(str(file.relative_to(root)).replace('\\', '/'))

    bridge = read('NativeHarnessBridge.json')
    lambda_bridge = read('LambdaHarnessBridge.json')
    profile_bridge = read('ProfileSnapshotBridge.json')
    if bridge.get('status') != 'PASS' or bridge.get('testedWorkbookSHA256') != tested_hash or bridge.get('releaseWorkbookSHA256') != logs.sha(root/'Release.xlsm'):
        errors.append('Native-only intermediate/tested bridge is missing or changed')
    if lambda_bridge.get('status') != 'PASS' or lambda_bridge.get('beforeWorkbookSHA256') != bridge.get('releaseWorkbookSHA256') or lambda_bridge.get('afterWorkbookSHA256') != lambda_hash:
        errors.append('Exact lambda assertion bridge is missing or changed')
    if not logs.bridge_chain(root) or profile_bridge.get('afterWorkbookSHA256') != release_hash:
        errors.append('Exact visualization diagram-selection bridge is missing or changed')
    for name in ('WrapperRemovalEvidence.json', 'SnapshotFixedUserData.json', 'SnapshotFixedSourceEquality.json',
                 'SnapshotFixedStructure.json', 'SnapshotFixedStandaloneHelp.json', 'FinalHelp/HelpEvidence.json'):
        artifact(name)
    result = {'status': 'PASS' if not errors else 'INCOMPLETE_OR_FAILED', 'sourceSHA256': release_hash,
              'testedSourceSHA256': tested_hash, 'groups': groups, 'artifacts': list(artifacts.values()),
              'errors': errors, 'scope': 'All requested standard, standalone ValidationPlan, geometry, normative, native, extension, plot and performance scenarios',
              'negativeProbesPreserved': True, 'sourceBridge': ['NativeHarnessBridge.json', 'LambdaHarnessBridge.json', 'ProfileSnapshotBridge.json']}
    (root/'FINAL_ACCEPTANCE_EVIDENCE.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps({'status': result['status'], 'groups': len(groups), 'artifacts': len(artifacts), 'errors': errors}))
    return 0 if not errors else 1


if __name__ == '__main__':
    raise SystemExit(main())
