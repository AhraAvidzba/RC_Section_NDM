"""ДЛЯ ТЕСТОВ: синтетические повреждённые/частичные логи; Excel/CAD/история не используются."""
import hashlib
import importlib.util
import json
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('closed_observer', HERE / 'TestObserveCompleted.py')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


def TestWrite(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False), encoding='utf-8')


def TestSelftest():
    checks = []
    with tempfile.TemporaryDirectory(prefix='NDMFullObserverSynthetic_') as temporary:
        root = Path(temporary)
        here = root / 'docs/regression/LaunchSynthetic'
        core = root / 'docs/regression/CoreSynthetic'
        request = dict(SourceSHA256='A' * 64, EvidenceDirectory='docs/regression/CoreSynthetic', ExpectedSuites=2, CoverageVersion=6)
        TestWrite(here / 'LaunchRequest.json', request)
        jobs = [dict(Group='Synthetic', Label=x, ExpectedSuites=1, Macro=['modTest.' + x]) for x in ('First', 'Second')]
        first = 'docs/regression/CoreSynthetic/First.txt'
        second = 'docs/regression/CoreSynthetic/Second.txt'
        bad = ('SOURCE_SHA256: ' + 'A' * 64 + '\nSUITE_STARTED: modTest.First; now\nFAIL spaced assertion\n'
               'RUNTIME ERROR: preserved\nSUITE_FINISHED: modTest.First; now\nWATCHDOG_COMPLETED: exit=1; source unchanged=True\n')
        (core / 'First.txt').parent.mkdir(parents=True, exist_ok=True)
        (core / 'First.txt').write_text(bad, encoding='utf-8')
        run = dict(SourceSHA256='A' * 64, CoverageVersion=6, CurrentJob='Second', Jobs=jobs,
                   CompletedJobs=[dict(Group='Synthetic', Label='First', Report=first, ExitCode=1, ExpectedSuites=1)])
        TestWrite(core / 'Run.json', run)
        observed = module.TestObserveCompleted(here, root, False)
        assert not observed['Records'][0]['FullSuiteCoverageProven']
        assert len(observed['Records'][0]['Failures']) == 2
        checks.append('FAIL-space and runtime with SUITE_FINISHED never prove PASS')
        (core / 'Second.txt').write_text('SOURCE_SHA256: ' + 'A' * 64 + '\nSUITE_STARTED: modTest.Second; now\nSUITE_FINISHED: modTest.Second; now\nWATCHDOG_COMPLETED: exit=0; source unchanged=True\n', encoding='utf-8')
        run['CurrentJob'] = 'Focused'
        run['CompletedJobs'].append(dict(Group='Synthetic', Label='Second', Report=second, ExitCode=0, ExpectedSuites=1))
        TestWrite(core / 'Run.json', run)
        observed = module.TestObserveCompleted(here, root, False)
        assert observed['Records'][0]['Failures'] and observed['Records'][1]['FullSuiteCoverageProven']
        assert not observed['FullRegressionPass'] and observed['Status'] == 'OBSERVATION_NOT_ACCEPTANCE'
        checks.append('Later clean group and stage switch preserve earlier FAIL; observer never grants full PASS')
        (core / 'First.txt').write_text(bad.replace('FAIL spaced assertion\n', '').replace('RUNTIME ERROR: preserved\n', '').replace('exit=1;', 'exit=0;'), encoding='utf-8')
        observed = module.TestObserveCompleted(here, root, False)
        assert observed['Records'][0]['ClosedPrimaryChanged'] and observed['Records'][0]['Failures']
        assert not observed['Records'][0]['FullSuiteCoverageProven']
        checks.append('Rewriting closed log cannot erase prior FAIL')
        TestWrite(core / 'Focused/Run.json', dict(SourceSHA256='A' * 64, Status='RUNNING', Records=[dict(ID='Independent', Kind='Python', Passed=True, Report='docs/regression/CoreSynthetic/Focused/Independent.json')]))
        TestWrite(core / 'Focused/Independent.json', dict(Passed=True, Unsupported=0, Failures=['ten expected numerical states'] * 10))
        observed = module.TestObserveCompleted(here, root, False)
        independent = next(x for x in observed['Records'] if x['Kind'] == 'FocusedPython')
        assert not independent['Failures'] and independent['JSONDiagnostic']['Passed']
        checks.append('Expected Independent.Failures diagnostics are not automatically classified as defects')
        before = (here / 'AllClosedJobsReview.json').read_bytes()
        run['SourceSHA256'] = 'B' * 64
        TestWrite(core / 'Run.json', run)
        try:
            module.TestObserveCompleted(here, root, False)
            raise RuntimeError('Source mismatch was accepted')
        except AssertionError:
            pass
        assert (here / 'AllClosedJobsReview.json').read_bytes() == before
        checks.append('Different run/source SHA rejected without destroying preserved review')
        run['SourceSHA256'] = 'A' * 64
        run['CompletedJobs'][0]['Report'] = 'docs/regression/OtherRun/Foreign.txt'
        TestWrite(core / 'Run.json', run)
        try:
            module.TestObserveCompleted(here, root, False)
            raise RuntimeError('Foreign historical report accepted')
        except AssertionError:
            pass
        checks.append('Foreign or historical report path rejected')
    result = dict(Status='PASS_SYNTHETIC_OBSERVER_ONLY', Checks=checks, Count=len(checks), RealLaunchPerformed=False,
                  ExcelProcessesStarted=0, CADProcessesStarted=0, HistoricalAcceptanceImported=False, FullRegressionPass=False,
                  ObserverSHA256=hashlib.sha256((HERE / 'TestObserveCompleted.py').read_bytes()).hexdigest().upper())
    (HERE / 'ObserverSyntheticSelftest.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(result, ensure_ascii=False))


if __name__ == '__main__':
    TestSelftest()
