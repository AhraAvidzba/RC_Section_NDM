"""Объединяет основные и целевые доказательства только для побайтно одной книги."""
import argparse
import hashlib
import importlib.util
import json
from pathlib import Path

ROOT=Path(__file__).resolve().parents[2]

def module(name,file):
    spec=importlib.util.spec_from_file_location(name,Path(__file__).with_name(file))
    value=importlib.util.module_from_spec(spec);spec.loader.exec_module(value);return value

def read(path):return json.loads(path.read_text(encoding='utf-8-sig'))
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest().upper()

def finalize(core,focused):
    main=module('main_suites','Summarize-FullOnOffAcceptance.py').summarize(core)
    if main['status']!='PASS':raise ValueError(main['errors'])
    run=read(focused/'Run.json')
    target=module('focused_suites','Verify-FocusedAcceptance.py').verify(focused,run)
    if target['Status']!='PASS':raise ValueError(target['Errors'])
    if main['sourceSHA256'].upper()!=target['SourceSHA256']:raise ValueError('Different workbook bytes')
    source=target['SourceSHA256']
    suites=main['observedSuites'];assertions=sum(x['passed'] for x in main['jobs'])
    # Исторические 80+36 проверены отдельными управляющими сценариями.
    if suites==80:
        plan=read(core/'SupplementaryPlan.json');extra=read(core/'SupplementaryRun.json')
        if extra['Status']!='COMPLETED_PENDING_LOG_REVIEW' or not extra['SourceUnchanged']:
            raise ValueError('Recipe supplements are incomplete')
        if plan['SourceSHA256']!=source or extra['SourceSHA256']!=source:
            raise ValueError('Recipe workbook provenance differs')
        if [x['Mode'] for x in extra['Runs']]!=['No','Yes']:raise ValueError('Both recipe modes required')
        logs=module('logs','Summarize-GeometryFull.py')
        for item in extra['Runs']:
            if item['ExitCode']!=0 or item['Macros']!=plan['Macros'] or len(item['Macros'])!=18:
                raise ValueError('Recipe suite coverage differs')
            path=ROOT/item['Report'];report=logs.inspect(path,source.lower(),18)
            if report['status']!='PASS' or [x['name'] for x in report['suites']]!=plan['Macros']:
                raise ValueError('Recipe log did not pass')
            body=path.read_text(encoding='utf-8-sig')
            if 'GLOBAL_MODE: General.DiagramExtension='+item['Mode']+';' not in body or 'WATCHDOG_COMPLETED: exit=0;' not in body:
                raise ValueError('Recipe mode or completion differs')
            suites+=len(report['suites']);assertions+=report['passed']
    expected = 116 if main.get('coverageDefinition') == 'HistoricalCoreSuitesOnly' else main['expectedSuites']
    if suites!=expected:raise ValueError(f'All {expected} core suites are required')
    independent=read(focused/'Independent.json')
    return {'Status':'PASS','SourceSHA256':source,'CoreSuites':suites,'CorePassedAssertions':assertions,
            'FocusedSingleCases':target['IndividualCases'],'FocusedPassedAssertions':target['PassedAssertions'],
            'IndependentIntegrals':independent['IndependentIntegrals'],
            'IndependentlyReviewedNumFailStates':independent['ReviewedNumFailStates'],
            'SameWorkbookBytes':True,'SingleInvocation':bool(main.get('includesFocusedAcceptance')),
            'CoreDirectory':str(core),'FocusedDirectory':str(focused),
            'FocusedRunSHA256':sha(focused/'Run.json'),'IndependentSHA256':sha(focused/'Independent.json'),
            'Meaning':'All current acceptance groups use exactly the same production workbook; invocation is recorded above.'}

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--core-directory',type=Path,required=True)
    parser.add_argument('--focused-directory',type=Path,required=True)
    args=parser.parse_args();result=finalize(args.core_directory,args.focused_directory)
    (args.focused_directory/'CombinedAcceptance.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False))
