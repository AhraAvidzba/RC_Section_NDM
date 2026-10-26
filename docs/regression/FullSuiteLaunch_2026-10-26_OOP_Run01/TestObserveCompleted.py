"""ДЛЯ ТЕСТОВ: читает завершённые журналы после смены активного этапа.

Дополняет наблюдение за текущим логом. Не меняет запущенные сценарии,
исходники, Run.json или официальные результаты приёмки.
"""
import hashlib
import json
import re
from datetime import datetime, timezone
from pathlib import Path


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestObserveCompleted():
    """Кэширует только неизменные завершённые журналы, сохраняя ошибки и runtime."""
    here = Path(__file__).resolve().parent
    root = here.parents[2]
    request = json.loads((here / 'LaunchRequest.json').read_text(encoding='utf-8-sig'))
    run = json.loads((root / request['EvidenceDirectory'] / 'Run.json').read_text(encoding='utf-8-sig'))
    assert run['SourceSHA256'] == request['SourceSHA256']
    target = here / 'AllClosedJobsReview.json'
    cached = json.loads(target.read_text(encoding='utf-8'))['Records'] if target.exists() else []
    records = []
    for job in run['CompletedJobs']:
        path = (root / job['Report']).resolve()
        assert path.is_relative_to((root / request['EvidenceDirectory']).resolve())
        raw = path.read_bytes()
        digest = hashlib.sha256(raw).hexdigest().upper()
        previous = next((x for x in cached if x['Report'] == job['Report'] and x['SHA256'] == digest), None)
        if previous:
            records.append(previous)
            continue
        text = raw.decode('utf-8-sig')
        failures = re.findall(r'^(?:FAIL:[^\r\n]*|RUNTIME ERROR[^\r\n]*|SCRIPT ERROR[^\r\n]*|WATCHDOG_FAILURE[^\r\n]*|RESTORE ERROR[^\r\n]*|CLEANUP ERROR[^\r\n]*)', text, re.M)
        runtimes = [x for x in failures if x.startswith(('RUNTIME ERROR', 'SCRIPT ERROR', 'WATCHDOG_FAILURE'))]
        finished = re.findall(r'^SUITE_FINISHED: ([^;\r\n]+)', text, re.M)
        records.append(dict(Report=job['Report'], SHA256=digest, ExitCode=job['ExitCode'],
                            ExpectedSuites=job['ExpectedSuites'], FinishedSuites=finished,
                            Failures=failures, RuntimeErrors=runtimes,
                            FullSuiteCoverageProven=len(finished) == job['ExpectedSuites'] and not runtimes))
    result = dict(Status='OBSERVATION_NOT_ACCEPTANCE', CheckedUTC=datetime.now(timezone.utc).isoformat(),
                  SourceSHA256=request['SourceSHA256'], CurrentJob=run.get('CurrentJob'), Records=records)
    temporary = target.with_suffix('.json.part')
    temporary.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    temporary.replace(target)
    print(json.dumps(dict(CompletedJobs=len(records), FailedJobs=[x['Report'] for x in records if x['ExitCode']],
                          Errors=[dict(Report=x['Report'], Failures=x['Failures']) for x in records if x['Failures']]), ensure_ascii=True))


if __name__ == '__main__':
    TestObserveCompleted()
