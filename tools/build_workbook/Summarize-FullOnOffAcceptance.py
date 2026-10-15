"""Проверяет полноту обоих режимов по журналам, хешам и завершению каждой suite."""
from __future__ import annotations
import argparse
import hashlib
import importlib.util
import json
import re
import subprocess
from datetime import datetime
from pathlib import Path

FULL = (
    'modTestGeometry.RunGeometryTests',
    'modTestMaterialDiagrams.RunMaterialDiagramTests',
    'modTestSectionSolver.RunSectionSolverTests',
    'modTestCapacitySolver.RunCapacitySolverTests',
    'modTestCrackWidth.RunCrackWidthTests',
    'modTestBatchCalculation.RunBatchCalculationTests',
    'modTestWorkbookInterface.RunWorkbookInterfaceTests',
    'modTestRegressionBaseline.RunRegressionBaselineTests',
)

def load_logs():
    path=Path(__file__).with_name('Summarize-GeometryFull.py')
    spec=importlib.util.spec_from_file_location('full_logs',path)
    module=importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

def summarize(directory):
    logs=load_logs()
    manifest_path=directory/'Run.json'
    manifest=json.loads(manifest_path.read_text(encoding='utf-8-sig'))
    source=manifest['SourceSHA256'].lower()
    jobs=[]
    errors=[]
    completed={(j['Group'],j['Mode']):j for j in manifest['CompletedJobs']}
    for job in manifest['Jobs']:
        path=directory/(job['Group']+'_'+job['Label']+'.txt')
        report=logs.inspect(path,source,job['ExpectedSuites'])
        expected=FULL if job['Group']=='Full' else tuple(job['Macro'])
        actual=tuple(s['name'] for s in report['suites'])
        text=path.read_text(encoding='utf-8-sig') if path.exists() else ''
        report['group']=job['Group']
        report['mode']=job['Mode']
        report['expectedMacros']=expected
        report['printedSuccessfulAssertions']=len(re.findall(r'(?m)^OK(?::|\|)',text))
        report['failures']=re.findall(r'(?m)^FAIL:.*$',text)
        if actual!=expected:
            report.setdefault('errors',[]).append('Suite names/order differ from the declared plan')
            report['status']='INCOMPLETE_OR_FAILED'
        expected_mode='GLOBAL_MODE: General.DiagramExtension='+job['Mode']+';'
        if expected_mode not in text:
            report.setdefault('errors',[]).append('Requested global mode is not confirmed')
            report['status']='INCOMPLETE_OR_FAILED'
        finish=completed.get((job['Group'],job['Mode']))
        if not finish or finish['ExitCode']!=0:
            report.setdefault('errors',[]).append('Successful driver completion is missing')
            report['status']='INCOMPLETE_OR_FAILED'
        if 'WATCHDOG_FAILURE:' in text or re.search(r'(?m)^(?:SCRIPT ERROR|CLEANUP ERROR|RUN_STOPPED):',text):
            report.setdefault('errors',[]).append('Runner failure or interruption')
            report['status']='INCOMPLETE_OR_FAILED'
        if report['status']!='PASS':errors.append(path.name)
        jobs.append(report)
    if manifest['Status']!='COMPLETED_PENDING_LOG_REVIEW':errors.append('Driver is not successfully complete')
    if manifest.get('SourceChanged'):errors.append('Frozen input changed')
    if manifest.get('RunnerChanged'):errors.append('Test runners changed during execution')
    for name,key in (('FullRunner.ps1','RunnerSHA256'),('ModeRunner.ps1','ModeRunnerSHA256')):
        snapshot=directory/name
        if not snapshot.is_file() or hashlib.sha256(snapshot.read_bytes()).hexdigest()!=manifest.get(key,'').lower():
            errors.append('Executed runner snapshot does not match: '+name)
    finished=manifest.get('FinishedUTC')
    if not finished or datetime.fromisoformat(finished)>datetime.fromisoformat(manifest['Deadline']):
        errors.append('Completion before the requested deadline is not confirmed')
    frozen=Path(manifest['FrozenWorkbook'])
    provenance='frozen file'
    if frozen.exists():
        input_bytes=frozen.read_bytes()
    else:
        # После очистки собственных временных копий точный вход восстанавливается из коммита.
        repo=Path(__file__).resolve().parents[2]
        try:
            relative=Path(manifest['SourceWorkbook']).relative_to(repo).as_posix()
            input_bytes=subprocess.check_output(['git','show',manifest['SourceCommit']+':'+relative],cwd=repo)
            provenance='committed workbook'
        except (ValueError,subprocess.CalledProcessError):
            input_bytes=b''
    if hashlib.sha256(input_bytes).hexdigest()!=source:
        errors.append('Frozen or committed workbook hash does not match the manifest')
    total_suites=sum(len(j['suites']) for j in jobs)
    expected_suites=sum(j['ExpectedSuites'] for j in manifest['Jobs'])
    return {
        'status':'PASS' if not errors else 'INCOMPLETE_OR_FAILED',
        'sourceCommit':manifest['SourceCommit'], 'sourceSHA256':source,
        'runnerSHA256':manifest['RunnerSHA256'],
        'modeRunnerSHA256':manifest.get('ModeRunnerSHA256'),
        'inputProvenance':provenance,
        'startedUTC':manifest['StartedUTC'], 'finishedUTC':finished,
        'deadline':manifest['Deadline'], 'expectedSuites':expected_suites,
        'completedSuites':sum(bool(s.get('finished')) for j in jobs for s in j['suites']),
        'observedSuites':total_suites,
        'printedSuccessfulAssertions':sum(j['printedSuccessfulAssertions'] for j in jobs),
        'failedAssertions':sum(len(j['failures']) for j in jobs),
        'errors':errors, 'jobs':jobs,
    }

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory',type=Path,required=True)
    parser.add_argument('--report',type=Path)
    parser.add_argument('--check',action='store_true')
    args=parser.parse_args()
    result=summarize(args.directory)
    output=args.report or args.directory/'Summary.json'
    output.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({k:v for k,v in result.items() if k!='jobs'},ensure_ascii=False))
    if args.check and result['status']!='PASS':raise SystemExit(1)

if __name__=='__main__':main()
