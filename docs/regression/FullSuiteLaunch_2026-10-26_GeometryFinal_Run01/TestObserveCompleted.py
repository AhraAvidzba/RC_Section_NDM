"""ДЛЯ ТЕСТОВ: сохраняет все закрытые primary, включая runtime, без присвоения полного PASS."""
import hashlib
import json
import re
import time
from datetime import datetime, timezone
from pathlib import Path

DEFAULT_FULL_MACROS = [
    'modTestGeometry.RunGeometryTests', 'modTestMaterialDiagrams.RunMaterialDiagramTests',
    'modTestSectionSolver.RunSectionSolverTests', 'modTestCapacitySolver.RunCapacitySolverTests',
    'modTestCrackWidth.RunCrackWidthTests', 'modTestBatchCalculation.RunBatchCalculationTests',
    'modTestWorkbookInterface.RunWorkbookInterfaceTests', 'modTestRegressionBaseline.RunRegressionBaselineTests',
]
ERROR = re.compile(r'^(?:FAIL(?::| )[^\r\n]*|RUNTIME ERROR[^\r\n]*|SCRIPT ERROR[^\r\n]*|WATCHDOG_FAILURE[^\r\n]*|RESTORE ERROR[^\r\n]*|CLEANUP ERROR[^\r\n]*)', re.M)


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestReadJSON(path):
    """Повторяет чтение того же manifest; transient запись не запускает новый процесс."""
    for attempt in range(4):
        try:
            return json.loads(path.read_text(encoding='utf-8-sig'))
        except (OSError, ValueError):
            if attempt == 3:
                raise
            time.sleep(.05)


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestReadText(raw):
    """Учитывает BOM UTF-16 и UTF-8; системный ANSI используется только без UTF-декодирования."""
    if raw.startswith((b'\xff\xfe', b'\xfe\xff')):
        return raw.decode('utf-16')
    try:
        return raw.decode('utf-8-sig')
    except UnicodeDecodeError:
        return raw.decode('cp1251')


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestObserveCompleted(here=None, root=None, emit=True):
    """Читает только текущий request; предыдущий FAIL сохраняется при смене этапа/правке закрытого лога."""
    here = Path(here) if here is not None else Path(__file__).resolve().parent
    root = Path(root) if root is not None else here.parents[2]
    request = TestReadJSON(here / 'LaunchRequest.json')
    directory = (root / request['EvidenceDirectory']).resolve()
    assert directory.is_relative_to((root / 'docs/regression').resolve()), 'Evidence outside docs/regression'
    run = TestReadJSON(directory / 'Run.json')
    assert run['SourceSHA256'].upper() == request['SourceSHA256'].upper(), 'Run/source SHA mismatch'
    target = here / 'AllClosedJobsReview.json'
    previous = TestReadJSON(target) if target.exists() else {}
    assert not previous or previous['SourceSHA256'].upper() == request['SourceSHA256'].upper(), 'Cross-source cache rejected'
    cached = previous.get('Records', [])
    records = []
    coverage_issues = []
    jobs = run.get('Jobs', [])
    if jobs:
        count = sum(int(job['ExpectedSuites']) for job in jobs)
        if count != int(request['ExpectedSuites']):
            coverage_issues.append(f'Planned core suite count {count} differs from requested {request["ExpectedSuites"]}')
        if run.get('CoverageVersion') != request.get('CoverageVersion'):
            coverage_issues.append('Core CoverageVersion differs from immutable request')
    closed = []
    for job in run.get('CompletedJobs', []):
        spec = next((x for x in jobs if x.get('Group') == job.get('Group') and x.get('Label') == job.get('Label')), {})
        macros = spec.get('Macro') or (DEFAULT_FULL_MACROS if job.get('Group') == 'Full' else [])
        closed.append(dict(job, Kind='CoreVBA', ExpectedMacros=macros))
    focused_path = directory / 'Focused/Run.json'
    focused = TestReadJSON(focused_path) if focused_path.exists() else None
    if focused:
        assert focused['SourceSHA256'].upper() == request['SourceSHA256'].upper(), 'Focused/source SHA mismatch'
        for record in focused.get('Records', []):
            if not record.get('Report'):
                continue
            item = dict(record)
            item['ExpectedMacros'] = [item['Macro']] if item.get('Kind') == 'VBA' and item.get('Macro') else []
            item['ExpectedSuites'] = 1 if item.get('Kind') == 'VBA' else None
            item['Kind'] = 'Focused' + str(item.get('Kind', ''))
            closed.append(item)
    for job in closed:
        path = (root / job['Report']).resolve()
        assert path.is_relative_to(directory), 'Closed report outside current evidence'
        raw = path.read_bytes()
        digest = hashlib.sha256(raw).hexdigest().upper()
        same_report = [x for x in cached if x['Report'] == job['Report']]
        old = next((x for x in same_report if x['SHA256'] == digest), None)
        if old:
            records.append(old)
            continue
        text = TestReadText(raw)
        failures = ERROR.findall(text)
        runtimes = [x for x in failures if x.startswith(('RUNTIME ERROR', 'SCRIPT ERROR', 'WATCHDOG_FAILURE'))]
        finished = re.findall(r'^SUITE_FINISHED: ([^;\r\n]+)', text, re.M)
        started = re.findall(r'^SUITE_STARTED: ([^;\r\n]+)', text, re.M)
        expected = job['ExpectedSuites']
        exact_order = finished == job['ExpectedMacros'] if job['ExpectedMacros'] else len(finished) == expected
        primary_source = re.search(r'^SOURCE_SHA256: ([^\r\n]+)', text, re.M)
        source_matches = bool(primary_source and primary_source[1].strip().upper() == request['SourceSHA256'].upper()) if job['Kind'] == 'CoreVBA' else None
        manifest_passed = job.get('Passed')
        json_note = None
        if path.suffix.lower() == '.json':
            payload = json.loads(text)
            json_note = dict(Passed=payload.get('Passed'), Status=payload.get('Status'), Unsupported=payload.get('Unsupported'),
                             Meaning='JSON Failures is diagnostic data; expected reviewed numerical rejections are not inferred as failures. Official verification remains authoritative.')
        prior_failures = [message for old_record in same_report for message in old_record.get('Failures', [])]
        failures = list(dict.fromkeys(prior_failures + failures))
        runtimes = list(dict.fromkeys([message for old_record in same_report for message in old_record.get('RuntimeErrors', [])] + runtimes))
        changed = bool(same_report and not old)
        complete = (expected is not None and len(finished) == expected and started == finished and exact_order
                    and not failures and not runtimes and job.get('ExitCode') == 0 and not changed
                    and (source_matches is not False) and manifest_passed is not False
                    and 'WATCHDOG_COMPLETED: exit=0;' in text)
        records.append(dict(Report=job['Report'], SHA256=digest, Kind=job['Kind'], ExitCode=job.get('ExitCode'),
                            ManifestPassed=manifest_passed, ExpectedSuites=expected, ExpectedMacros=job['ExpectedMacros'],
                            StartedSuites=started, FinishedSuites=finished, Failures=failures, RuntimeErrors=runtimes,
                            PrimarySourceMatches=source_matches, ClosedPrimaryChanged=changed,
                            PreviousSHA256=[x['SHA256'] for x in same_report if x['SHA256'] != digest],
                            FullSuiteCoverageProven=bool(complete), JSONDiagnostic=json_note))
    current_reports = {x['Report'] for x in records}
    records += [x for x in cached if x['Report'] not in current_reports]
    result = dict(Status='OBSERVATION_NOT_ACCEPTANCE', CheckedUTC=datetime.now(timezone.utc).isoformat(),
                  SourceSHA256=request['SourceSHA256'], CurrentJob=run.get('CurrentJob'), CoverageIssues=coverage_issues,
                  RequestedCoreSuites=request['ExpectedSuites'], CoverageVersion=request.get('CoverageVersion'),
                  FocusedStatus=focused.get('Status') if focused else None, Records=records,
                  FullRegressionPass=False, OfficialPrimaryReviewStillRequired=True)
    temporary = target.with_suffix('.json.part')
    temporary.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    temporary.replace(target)
    if emit:
        print(json.dumps(dict(Status=result['Status'], ClosedRecords=len(records), CoverageIssues=coverage_issues,
                              FailedRecords=[x['Report'] for x in records if x['Failures'] or x['ExitCode'] not in (None, 0) or x['ManifestPassed'] is False or x['ClosedPrimaryChanged']],
                              Errors=[dict(Report=x['Report'], Failures=x['Failures']) for x in records if x['Failures']]), ensure_ascii=True))
    return result


if __name__ == '__main__':
    TestObserveCompleted()
