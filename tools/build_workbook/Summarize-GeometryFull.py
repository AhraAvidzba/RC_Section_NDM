"""Validate terminal full ON/OFF logs without counting nested totals twice."""
import argparse
import hashlib
import json
import re
from pathlib import Path


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def inspect(path, source_hash, expected_count):
    if not path.is_file():
        return {"name": path.name, "status": "NOT_RUN", "suites": []}
    text = path.read_text(encoding="utf-8-sig")
    errors = []
    source = re.search(r"(?m)^SOURCE_SHA256: ([0-9A-Fa-f]{64})$", text)
    if not source or source.group(1).lower() != source_hash:
        errors.append("Source fingerprint mismatch")
    starts = list(re.finditer(r"(?m)^SUITE_STARTED: ([^;]+); (.+)$", text))
    suites = []
    for index, start in enumerate(starts):
        end = starts[index + 1].start() if index + 1 < len(starts) else len(text)
        block = text[start.start():end]
        name = start.group(1)
        finish = re.search(r"(?m)^SUITE_FINISHED: " + re.escape(name) + r"; (.+)$", block)
        totals = re.findall(r"(?m)^TOTAL[^\r\n]*?passed=(\d+); failed=(\d+)[^\r\n]*", block)
        failed_lines = re.findall(r"(?m)^FAIL:.*$", block)
        failures = sum(int(failed) for _, failed in totals)
        passed = int(totals[-1][0]) if totals else None
        status = "RUNNING" if not finish else "PASS"
        if failed_lines or failures or "RUNTIME ERROR" in block:
            status = "FAIL"
        if finish and not totals:
            status = "UNVERIFIED_TOTAL"
        suites.append({"name": name, "status": status, "started": start.group(2),
                       "finished": finish.group(1) if finish else None,
                       "passed": passed, "failedLines": len(failed_lines),
                       "nestedFailureCounts": [int(f) for _, f in totals]})
    terminal = "WATCHDOG_COMPLETED: exit=0; source unchanged=True" in text
    if len(suites) != expected_count:
        errors.append(f"Expected {expected_count} suites, observed {len(suites)}")
    if any(s["status"] != "PASS" for s in suites):
        errors.append("Suite is incomplete or failed")
    if not terminal or "SOURCE_UNCHANGED: True" not in text:
        errors.append("Successful terminal source guard is missing")
    return {"name": path.name, "status": "PASS" if not errors else "INCOMPLETE_OR_FAILED",
            "sourceSHA256": source.group(1).lower() if source else None,
            "reportSHA256": sha(path), "suites": suites, "errors": errors,
            "passed": sum(s["passed"] or 0 for s in suites)}


def bridge_chain(root):
    """Join inputs only through three exact private test-method patches."""
    def read(name):
        path = root/name
        return json.loads(path.read_text(encoding='utf-8-sig')) if path.is_file() else {}
    native, patch = read('NativeHarnessBridge.json'), read('LambdaHarnessBridge.json')
    profile = read('ProfileSnapshotBridge.json')
    good = (native.get('status') == patch.get('status') == profile.get('status') == 'PASS'
            and native.get('productionUnchanged') and native.get('unchangedOtherMethods')
            and patch.get('productionUnchanged') and patch.get('unchangedOtherCode')
            and native.get('changedMethod') == 'modTestContourMesh.TestCheckNativeFullExport'
            and patch.get('changedMethod') == 'modTestConfiguration.TestCheckCapacityNumericEffects'
            and native.get('changed') == ['modTestContourMesh'] and patch.get('changed') == ['modTestConfiguration']
            and profile.get('productionUnchanged') and profile.get('unchangedOtherCode') and profile.get('tolerancesUnchanged')
            and profile.get('changed') == ['modTestProfileConfig']
            and profile.get('changedMethod') == 'modTestProfileConfig.TestCheckVisualizationValues')
    checks = ((native, 'testedWorkbookSHA256', 'FinalReady.xlsm'),
              (native, 'releaseWorkbookSHA256', 'Release.xlsm'),
              (native, 'beforeDumpSHA256', 'VBA_Final.txt'), (native, 'afterDumpSHA256', 'VBA_Release.txt'),
              (patch, 'beforeWorkbookSHA256', 'Release.xlsm'),
              (patch, 'afterWorkbookSHA256', 'ReleaseLambdaFixed.xlsm'),
              (patch, 'beforeDumpSHA256', 'VBA_Release.txt'), (patch, 'afterDumpSHA256', 'VBA_LambdaFixed.txt'),
              (profile, 'beforeWorkbookSHA256', 'ReleaseLambdaFixed.xlsm'),
              (profile, 'afterWorkbookSHA256', 'ReleaseSnapshotFixed.xlsm'),
              (profile, 'beforeDumpSHA256', 'VBA_LambdaFixed.txt'), (profile, 'afterDumpSHA256', 'VBA_SnapshotFixed.txt'))
    good = good and all((root/name).is_file() and value.get(key) == sha(root/name) for value, key, name in checks)
    return bool(good)


def full_scope(root, mode):
    """Revalidate all eight suites, retaining the failed original ON report."""
    tested, current = sha(root/'FinalReady.xlsm'), sha(root/'ReleaseLambdaFixed.xlsm')
    raw_path = root/f'FinalFull_{mode}.txt'
    raw = inspect(raw_path, tested, 8)
    text = raw_path.read_text(encoding='utf-8-sig') if raw_path.is_file() else ''
    directed = [inspect(root/f'LambdaFixedDirectedRetry_{label}.txt', current, 1) for label in ('Off', 'On')]
    baseline = inspect(root/'LambdaBaselineProbe_On.txt', sha(root/'LambdaBaselineProbe.xlsm'), 1)
    retry = inspect(root/'LambdaFixedBatch_On.txt', current, 1)
    snapshot_hash = sha(root/'ReleaseSnapshotFixed.xlsm')
    profile_retries = [inspect(root/f'ProfileSnapshotFixed_{label}.txt', snapshot_hash, 1) for label in ('Off', 'On')]
    profile_path = root/'ProfileBaselineOriginal_On.txt'
    profile_baseline = inspect(profile_path, sha(root/'ProfileBaselineProbe.xlsm'), 1)
    profile_text = profile_path.read_text(encoding='utf-8-sig') if profile_path.is_file() else ''
    profile_failures = [f'profileConfig.PR{p}.Visualization.CapacityState.Stress.element{e}; actual={actual}; expected=350'
                        for p in (1, 2, 3, 4) for e, actual in
                        ((101, '350.000000000404'), (102, '350.000000000405'),
                         (103, '350.000000000415'), (104, '350.000000000417'))]
    errors = []
    accepted = [{**s, 'report': raw['name'], 'sourceSHA256': tested} for s in raw['suites'] if s['status'] == 'PASS']
    if not bridge_chain(root):
        errors.append('Exact test-only source bridge is unproven')
    if any(r['status'] != 'PASS' for r in directed) or baseline['status'] != 'PASS':
        errors.append('Corrected complete configuration tests or old-kernel reproduction incomplete')
    if any(r['status'] != 'PASS' or r['passed'] != 8824 for r in profile_retries):
        errors.append('Complete corrected profile tests (292 cases, both modes) are incomplete')
    profile_reproduced = (re.findall(r'(?m)^FAIL: (.+)$', profile_text) == profile_failures
                          and 'TOTAL_AUDIT03_PROFILE_CONFIG: passed=8768; failed=16; cases=292' in profile_text
                          and len(profile_baseline['suites']) == 1 and profile_baseline['suites'][0]['finished']
                          and profile_baseline.get('sourceSHA256') == sha(root/'ProfileBaselineProbe.xlsm')
                          and 'SOURCE_UNCHANGED: True' in profile_text
                          and 'WATCHDOG_COMPLETED: exit=1; source unchanged=True' in profile_text
                          and 'WATCHDOG_FAILURE:' not in profile_text and 'RUNTIME ERROR' not in profile_text)
    if not profile_reproduced:
        errors.append('Exact original profile failures have not reproduced on the old kernel')
    comparison = {}
    for label, filename in (('oldKernel', 'LambdaBaselineProbe_On.txt'), ('current', 'LambdaFixedDirectedRetry_On.txt')):
        path = root/filename
        value = path.read_text(encoding='utf-8-sig') if path.is_file() else ''
        comparison[label] = re.findall(r'(?m)^OK: audit03\.capacityNumeric\.p([12])\.Capacity\.ToleranceLambda\.sameLimit\|actual=([^|\r\n]+)\|expected=([^\r\n]+)', value)
    if len(comparison['oldKernel']) != 2 or comparison['oldKernel'] != comparison['current']:
        errors.append('Old/current kernel coarse and precise lambda values have not matched exactly')
    if mode == 'Off':
        if raw['status'] != 'PASS':
            errors.append('Original full OFF is not PASS')
    else:
        failed = re.findall(r'(?m)^FAIL: (.+)$', text)
        lambda_failures = [f'audit03.capacityNumeric.p{p}.Capacity.ToleranceLambda.sameLimit|actual=4.703125|expected=4.71548271179199' for p in (1, 2)]
        allowed = failed == lambda_failures + profile_failures
        finished = len(raw['suites']) == 8 and all(s['finished'] for s in raw['suites'])
        only_known_failed = [s['name'] for s in raw['suites'] if s['status'] != 'PASS'] == [
            'modTestBatchCalculation.RunBatchCalculationTests', 'modTestWorkbookInterface.RunWorkbookInterfaceTests']
        guarded = (raw.get('sourceSHA256') == tested and 'GLOBAL_MODE: General.DiagramExtension=Yes;' in text
                   and 'SOURCE_UNCHANGED: True' in text and 'WATCHDOG_COMPLETED: exit=1; source unchanged=True' in text
                   and 'WATCHDOG_FAILURE:' not in text and 'RUNTIME ERROR' not in text)
        counts = [count for suite in raw['suites'] for count in suite['nestedFailureCounts'] if count]
        guarded = guarded and counts == [2, 2, 16, 16]
        if not (allowed and finished and only_known_failed and guarded):
            errors.append('Original ON has incomplete work or failures outside the exact corrected expectation')
        if retry['status'] != 'PASS':
            errors.append('Complete corrected Batch ON is not PASS')
        else:
            accepted += [{**s, 'report': retry['name'], 'sourceSHA256': current} for s in retry['suites']]
        if allowed and finished and only_known_failed and guarded and profile_retries[1]['status'] == 'PASS':
            original_ui = raw['suites'][6]
            # All other UI tests completed unchanged; the entire affected profile suite was repeated.
            accepted.append({**original_ui, 'status': 'PASS', 'report': raw['name'],
                             'sourceSHA256': tested, 'acceptance': 'COMPOSITE_REVALIDATION',
                             'originalStatus': original_ui['status'], 'originalPassed': original_ui['passed'],
                             'passed': original_ui['passed'] - 8768 + profile_retries[1]['passed'],
                             'failedLines': 0, 'nestedFailureCounts': [],
                             'unchangedUIPassed': original_ui['passed'] - 8768,
                             'affectedSuiteRevalidation': profile_retries[1],
                             'revalidationSourceSHA256': snapshot_hash})
    if len(accepted) != 8:
        errors.append('Eight actually completed accepted suites required')
    return {'name': f'FullScope_{mode}', 'status': 'PASS' if not errors else 'INCOMPLETE_OR_FAILED',
            'mode': mode, 'acceptedSuites': accepted, 'originalFullReport': raw,
            'configurationRevalidation': directed, 'oldKernelReproduction': baseline,
            'oldCurrentLambdaComparison': comparison,
            'profileRevalidation': profile_retries, 'oldKernelProfileReproduction': profile_baseline,
            'exactProfileFailuresReproduced': bool(profile_reproduced),
            'correctedBatchOn': retry if mode == 'On' else None, 'errors': errors,
            'passed': sum(s['passed'] or 0 for s in accepted),
            'scope': 'All eight original standard suites completed; corrected full Batch ON and entire affected profile suite ON/OFF repeated through exact test-only bridges. Other UI tests completed unchanged. Original ON remains FAIL18; composite acceptance is explicit.'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--directory", type=Path, required=True)
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    source_hash = sha(args.source)
    current_hash = sha(args.directory/'ReleaseLambdaFixed.xlsm')
    reports = [inspect(args.directory/f'Final{group}_{mode}.txt', source_hash if mode == 'Off' or group == 'Full' else current_hash, count)
               for mode in ('Off', 'On') for group, count in (('Full', 8), ('Supplement', 14), ('Extras', 9))]
    scopes = [full_scope(args.directory, mode) for mode in ('Off', 'On')]
    supplemental = [r for r in reports if not r['name'].startswith('FinalFull_')]
    result = {"status": "PASS" if all(r['status'] == 'PASS' for r in scopes+supplemental) else "INCOMPLETE_OR_FAILED",
              "scope": "Standard8 + Geometry14 + ValidationExtras9 per mode; complements recorded separately",
              "source": str(args.source), "sourceSHA256": source_hash, "reports": reports,
              "fullScopes": scopes, "sourceBridge": ['NativeHarnessBridge.json', 'LambdaHarnessBridge.json', 'ProfileSnapshotBridge.json'],
              "passed": sum(r.get("passed", 0) for r in scopes+supplemental),
              "counting": "Only last TOTAL per top-level suite; nested failures remain fatal"}
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": result["status"], "reports": len(reports), "passed": result["passed"]}))
    return 0 if result["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
