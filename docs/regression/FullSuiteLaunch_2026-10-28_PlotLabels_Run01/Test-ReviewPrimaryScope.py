"""ДЛЯ ТЕСТОВ: полностью читает первичные отчёты текущего полного набора; не запускает Excel и не присваивает PASS."""
from pathlib import Path
import argparse,hashlib,importlib.util,json,re,subprocess
from datetime import datetime,timezone
parser=argparse.ArgumentParser();parser.add_argument('--launch-directory',type=Path);args=parser.parse_args()
here=(args.launch_directory or Path(__file__).resolve().parent).resolve();root=here.parents[2]
def read(p):return json.loads(p.read_text('utf-8-sig'))
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest().upper()
request=read(here/'LaunchRequest.json');launch=read(here/'FullLaunch.json')
assert sha(here/'LaunchRequest.json')==launch['RequestSHA256']
directory=(root/request['EvidenceDirectory']).resolve();assert directory.is_relative_to(root/'docs/regression')
run=read(directory/'Run.json');assert run['SourceSHA256']==request['SourceSHA256']
spec=importlib.util.spec_from_file_location('current_observer',here/'TestObserveCompleted.py');observer=importlib.util.module_from_spec(spec);spec.loader.exec_module(observer)
primary=[];issues=[];failures=[];identities={};all_cases={};core_macros=0;closed_count=0
def identity(name,pid,ticks,provenance):
    assert name in ('EXCEL','acad','powershell') and int(pid)>0 and int(ticks)>0
    key=(name,int(pid),str(ticks));entry=identities.setdefault(key,{'Name':name,'PID':int(pid),'StartTicksUTC':str(ticks),'Provenance':[]})
    if provenance not in entry['Provenance']:entry['Provenance'].append(provenance)
def inspect(path,role,expected_source=None,macros=None,closed=True):
    assert path.resolve().is_relative_to(directory)
    raw=path.read_bytes();text=observer.TestReadText(raw)
    relative=str(path.relative_to(root)).replace('\\','/')
    bad=observer.ERROR.findall(text)
    source=re.findall(r'(?m)^SOURCE_SHA256: ([0-9A-F]+)\r?$',text)
    source_matches=source==[expected_source] if expected_source else None
    started=re.findall(r'^SUITE_STARTED: ([^;\r\n]+)',text,re.M)
    finished=re.findall(r'^SUITE_FINISHED: ([^;\r\n]+)',text,re.M)
    if macros is not None and closed:
        if started!=macros or finished!=macros:issues.append(relative+': macro names/order/count differ')
        if source_matches is False:issues.append(relative+': source SHA differs')
        if 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' not in text:issues.append(relative+': missing successful watchdog/source guard')
    for pid,ticks in re.findall(r'TEST_EXCEL_PROCESS: id=(\d+); startTicks=(\d+)',text):identity('EXCEL',pid,ticks,relative)
    failures.extend({'Path':relative,'Message':line} for line in bad)
    entry={'Path':relative,'Role':role,'Bytes':len(raw),'SHA256':hashlib.sha256(raw).hexdigest().upper(),'AllBytesRead':True,'Closed':closed,'SourceMatches':source_matches,'SourceSHA256':source,'Started':started,'Finished':finished,'Errors':bad,'TotalLines':re.findall(r'(?m)^TOTAL[^\r\n]*',text)}
    primary.append(entry);return text
for name in ('DriverRun.json','HostRun.json'):
    value=read(here/name);identity('powershell',value['PID'],value['StartTicks'],str((here/name).relative_to(root)).replace('\\','/'))
completed={(r['Group'],r['Mode']):r for r in run.get('CompletedJobs',[])}
assert len(run['Jobs'])==12 and sum(j['ExpectedSuites'] for j in run['Jobs'])==182
for job in run['Jobs']:
    path=directory/(job['Group']+'_'+job['Label']+'.txt');done=completed.get((job['Group'],job['Mode']))
    if not path.exists():continue
    expected=job['Macro'] or observer.DEFAULT_FULL_MACROS
    text=inspect(path,'Core',request['SourceSHA256'],expected,bool(done))
    if done:
        closed_count+=1;core_macros+=len(re.findall(r'^SUITE_FINISHED:',text,re.M))
        if done.get('ExitCode')!=0:issues.append(path.name+': unsuccessful worker exit')
        if job['Group']=='Full' and 'TOTAL_WORKBOOK_UI:' not in text:issues.append(path.name+': complete UI terminal total missing')
focused_path=directory/'Focused/Run.json';focused=read(focused_path) if focused_path.exists() else None
focused_records=focused.get('Records',[]) if focused else []
if focused:
    assert focused['SourceSHA256']==request['SourceSHA256']
    prep=read(directory/'Focused/Probe/Preparation.json');assert prep['SourceSHA256']==request['SourceSHA256'] and prep['Compiled'] and not prep['ProductionChanged']
    assert sha(directory/'Focused/Probe/RC_Section_NDM.xlsm')==prep['ProbeWorkbookSHA256']
    for record in focused_records:
        path=(root/record['Report']).resolve();assert path.is_relative_to(directory)
        if not record.get('Passed'):issues.append('Focused '+record['ID']+': Passed is false')
        if path.suffix=='.txt':
            text=inspect(path,'Focused',prep['ProbeWorkbookSHA256'],[record['Macro']])
            if sha(path)!=record['ReportSHA256']:issues.append(record['ID']+': recorded report SHA differs')
            if record['Macro']=='modTestBatchCalculation.TestRunFocusedAcceptanceCases':
                manifest=directory/'Focused/Probe'/(record['Argument1']+'_'+record['Argument2']+'_Manifest.txt')
                wanted=re.findall(r'^FOCUSED_INPUT\|id=([^|\r\n]+)',manifest.read_text('utf-8-sig'),re.M)
                actual=re.findall(r'^FOCUSED_TIME\|[^\r\n]*?\|id=([^|\r\n]+)',text,re.M)
                if actual!=wanted or len(actual)!=len(set(actual)):issues.append(record['ID']+': actual individual IDs differ from input manifest')
                all_cases[record['ID']]=actual
        else:
            inspect(path,'FocusedJSON');payload=read(path)
            if payload.get('Passed') is not True:issues.append(record['ID']+': primary JSON is not passed')
            if record['ID']=='Independent':
                if payload.get('SourceSHA256')!=request['SourceSHA256'] or payload.get('Unsupported'):issues.append('Independent source/unsupported mismatch')
                if payload.get('ModelCount')!=5 or payload.get('ReviewedNumFailStates')!=10:issues.append('Independent model/numerical review coverage differs')
                if len(payload.get('Failures',[]))!=10 or any(not r.get('Passed') for r in payload.get('Failures',[])):issues.append('Expected numerical rejections lack individual proof')
            if record['ID']=='NativeCAD':
                if {r['Macro'] for r in payload['Results']}!={'RunRealAutoCADFarOriginTests','RunRealAutoCADFarCapacityTests'}:issues.append('Native CAD composition differs')
                for result in payload['Results']:
                    folder=path.parent/result['Macro'];invocation=read(folder/'NativeInvocation.json')
                    if invocation['SourceSHA256']!=request['SourceSHA256'] or not invocation['Passed'] or not invocation['SourceUnchanged']:issues.append('Native CAD invocation/source differs')
                    child=folder/'NativeCAD.txt';inspect(child,'NativeCADChild')
                    if sha(child)!=result['ReportSHA256']:issues.append('Native CAD child SHA differs')
for path in directory.rglob('*Ownership.json'):
    value=read(path);relative=str(path.relative_to(root)).replace('\\','/')
    if 'OwnedPID' in value:
        name='acad' if path.name=='NativeCADOwnership.json' else value.get('ProcessName')
        if name in ('EXCEL','acad'):identity(name,value['OwnedPID'],value['StartTicks'],relative)
    elif value.get('VerifiedNewOwner') is True and value.get('ProcessName')=='EXCEL':identity('EXCEL',value['PID'],value['StartTicksUTC'],relative)
stdio=here/'FullLaunch.stdout.log'
for pid,ticks in re.findall(r'OWNED_EXCEL: id=(\d+); utcStartTicks=(\d+)',observer.TestReadText(stdio.read_bytes())):identity('EXCEL',pid,ticks,str(stdio.relative_to(root)).replace('\\','/'))
input_rows=request['ProtectedInputs']+request['SourceComponents']+request['ParticipatingScripts']
changed=[r['Path'] for r in input_rows if not (root/r['Path']).is_file() or sha(root/r['Path'])!=r['SHA256']]
head_same=subprocess.check_output(['git','rev-parse','HEAD'],cwd=root,text=True).strip()==request['SourceCommit']
ended=bool((here/'Completion.json').exists() and run.get('FinishedUTC'))
if ended:
    if closed_count!=12 or core_macros!=182:issues.append('Entire 12 core/182 macro composition is incomplete')
    if len(focused_records)!=26 or sum(len(v) for v in all_cases.values())!=122:issues.append('Entire 26 Focused/122 individual composition is incomplete')
    if changed or not head_same:issues.append('Frozen inputs/HEAD changed')
result={'Status':'ENTIRE_TERMINAL_PRIMARY_READ_NOT_ACCEPTANCE' if ended else 'PARTIAL_PRIMARY_READ_NOT_ACCEPTANCE','CheckedUTC':datetime.now(timezone.utc).isoformat(),'SourceSHA256':request['SourceSHA256'],'SourceCommit':request['SourceCommit'],'PrimaryLogsReviewed':ended and closed_count==12 and len(focused_records)==26,'CoreReportsRead':closed_count,'CoreMacrosReviewed':core_macros,'FocusedRecordCount':len(focused_records),'FocusedSingleCasesReviewed':sum(len(v) for v in all_cases.values()),'FocusedCaseIDs':all_cases,'Primary':primary,'Identities':list(identities.values()),'InventoryErrors':issues,'Failures':failures,'InputsChecked':len(input_rows),'ChangedInputs':changed,'HEADUnchanged':head_same,'FullAcceptanceClaim':False}
target=here/('TerminalScopeReview.json' if ended else 'TerminalScopeReviewPartial.json');target.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({k:v for k,v in result.items() if k not in ('Primary','Identities','FocusedCaseIDs')},ensure_ascii=False))
