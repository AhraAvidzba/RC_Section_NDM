"""ДЛЯ ТЕСТОВ: перечитывает закрытые core-журналы без изменения исходного прогона."""
from pathlib import Path
from datetime import datetime
import hashlib
import importlib.util
import json
import re
import sys


def TestReviewClosedPrimary():
    """Сохраняет первичный разбор каждого закрытого job; общий PASS здесь не устанавливается."""
    root = Path.cwd()
    launch = root / 'docs/regression/FullSuiteLaunch_2026-10-26_OOP_Run01'
    request = json.loads((launch / 'LaunchRequest.json').read_text(encoding='utf-8-sig'))
    directory = (root / request['EvidenceDirectory']).resolve()
    run = json.loads((directory / 'Run.json').read_text(encoding='utf-8-sig'))
    assert run['SourceSHA256'] == request['SourceSHA256']
    observation = json.loads((launch / 'AllClosedJobsReview.json').read_text(encoding='utf-8-sig'))
    known = {record['Report']: record for record in observation['Records']}
    spec = importlib.util.spec_from_file_location('closed_primary', root / 'tools/build_workbook/Summarize-GeometryFull.py')
    verifier = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(verifier)
    records = []
    for job in run['CompletedJobs']:
        path = (root / job['Report']).resolve()
        assert path.is_relative_to(directory)
        raw = path.read_bytes()
        text = raw.decode('utf-8-sig')
        digest = hashlib.sha256(raw).hexdigest().upper()
        assert job['Report'] in known and digest == known[job['Report']]['SHA256']
        owner = re.search(r'TEST_EXCEL_PROCESS: id=(\d+); startTicks=(\d+)', text)
        inspection = verifier.inspect(path, request['SourceSHA256'].lower(), job['ExpectedSuites'])
        mode_ok = ('GLOBAL_MODE: General.DiagramExtension=' + job['Mode'] + ';') in text
        passed = job['ExitCode'] == 0 and inspection.get('status') == 'PASS' and mode_ok and owner is not None
        records.append({'Group': job['Group'], 'Mode': job['Mode'], 'Report': job['Report'],
                        'SHA256': digest, 'Bytes': len(raw), 'ExitCode': job['ExitCode'],
                        'ExpectedSuites': job['ExpectedSuites'], 'Inspection': inspection,
                        'GlobalModeMatches': mode_ok, 'PassedClosedGroup': passed,
                        'OwnedExcel': {'PID': int(owner[1]), 'StartTicks': owner[2]} if owner else None,
                        'TotalsOfInterest': re.findall(r'^TOTAL(?:_WORKBOOK_UI|_BATCH|_REGRESSION_BASELINE)[^\r\n]*', text, re.M),
                        'EndMarkers': re.findall(r'^(?:SUITE_FINISHED:|SOURCE_UNCHANGED:|WATCHDOG_COMPLETED:)[^\r\n]*', text, re.M)})
    result = {'At': datetime.now().astimezone().isoformat(),
              'Status': 'CLOSED_CORE_PRIMARY_REVIEW_NOT_FULL_ACCEPTANCE',
              'SourceSHA256': request['SourceSHA256'], 'CurrentJob': run.get('CurrentJob'),
              'ClosedJobs': len(records), 'ClosedSuites': sum(len(r['Inspection'].get('suites', [])) for r in records),
              'AllClosedGroupsPassed': all(r['PassedClosedGroup'] for r in records),
              'Records': records, 'FullPass': False}
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    out = root / ('docs/regression/OOPLayers_2026-10-26_Run01/Observations/ClosedPrimaryReview_' + stamp + '.json')
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'Proof': str(out.relative_to(root)), 'ClosedJobs': len(records),
                      'ClosedSuites': result['ClosedSuites'], 'AllClosedGroupsPassed': result['AllClosedGroupsPassed'],
                      'Groups': [{'Group': r['Group'], 'Mode': r['Mode'], 'Suites': len(r['Inspection'].get('suites', [])),
                                  'Status': r['Inspection'].get('status'), 'ModeMatches': r['GlobalModeMatches'],
                                  'Owner': r['OwnedExcel']} for r in records]}, ensure_ascii=True))
    return 0 if result['AllClosedGroupsPassed'] else 1


if __name__ == '__main__':
    sys.exit(TestReviewClosedPrimary())