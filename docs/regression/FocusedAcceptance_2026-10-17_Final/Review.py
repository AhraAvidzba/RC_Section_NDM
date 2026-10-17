"""Повторная проверка принятого входа, полноты, отрицательных контролей и времени."""
from pathlib import Path
import copy,datetime,hashlib,importlib.util,json,re,statistics
D=Path(__file__).resolve().parent
ROOT=D.parents[2]
def read(p):return json.loads(p.read_text(encoding='utf-8-sig'))
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest().upper()
spec=importlib.util.spec_from_file_location('focused',ROOT/'tools/build_workbook/Verify-FocusedAcceptance.py')
m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
run=read(D/'Run.json');accepted=m.verify(D,run)
assert accepted['Status']=='PASS',accepted['Errors']
controls=[]
for name in ('missing_load_group','missing_native','missing_independent','wrong_mode','wrong_source','missing_finish','late_finish','failed_independent_exit'):
    r=copy.deepcopy(run)
    if name=='missing_load_group':r['Records']=[x for x in r['Records'] if x['ID']!='SavedWorkbook_Paths_Yes']
    if name=='missing_native':r['Records']=[x for x in r['Records'] if x['ID']!='NativeCAD']
    if name=='missing_independent':r['Records']=[x for x in r['Records'] if x['ID']!='Independent']
    if name=='wrong_mode':next(x for x in r['Records'] if x['ID']=='Circle_Generated_Yes')['Mode']='No'
    if name=='wrong_source':r['SourceSHA256']='0'*64
    if name=='missing_finish':del r['FinishedUTC']
    if name=='late_finish':r['FinishedUTC']=(datetime.datetime.fromisoformat(r['DeadlineUTC'])+datetime.timedelta(seconds=1)).isoformat()
    if name=='failed_independent_exit':next(x for x in r['Records'] if x['ID']=='Independent')['Passed']=False
    check=m.verify(D,r)
    assert check['Status']!='PASS',name
    controls.append({'Name':name,'Rejected':True,'Errors':check['Errors']})
timing=[];groups=[];batches={};specials={}
for record,lines in m.reports(D,run):
    times=[m.fields(x) for x in lines if x.startswith('FOCUSED_TIME|')]
    if times:
        seconds=[float(x['elapsedSec']) for x in times]
        timing.extend({'Group':record['ID'],'ID':x['id'],'Seconds':float(x['elapsedSec'])} for x in times)
        groups.append({'Group':record['ID'],'Cases':len(times),'MaxSeconds':max(seconds)})
    for line in lines:
        if line.startswith('CORE_BATCH_TIME|'):batches[record['Mode']]=float(m.fields(line)['seconds'])
        if line.startswith('TOTAL_FOCUSED_SPECIAL:'):
            specials[record['ID']]=float(re.search(r'elapsedSec=([^;]+)',line)[1])
old=read(ROOT/'docs/regression/Performance/CoreRegression_2026-10-15/After/InputManifest.json')
current=read(D/'Probe/InputManifest.json');matching=[]
for item in current['Groups']:
    before=next(x for x in old['Groups'] if (x['Shape'],x['Group'])==(item['Shape'],item['Group']))
    assert before['ReportSHA256']==item['ReportSHA256']
    matching.append({'Shape':item['Shape'],'Group':item['Group'],'Exact':True,'SHA256':item['ReportSHA256']})
proof=read(D/'Independent.json')
result={'Status':'PASS','SourceSHA256':run['SourceSHA256'],'IndividualCases':len(timing),
        'MaxIndividualSeconds':max(x['Seconds'] for x in timing),'MedianIndividualSeconds':statistics.median(x['Seconds'] for x in timing),
        'BatchSeconds':batches,'SpecialSeconds':specials,'Groups':groups,'ExactHistoricalInputs':matching,
        'NegativeControls':controls,'IndependentIntegrals':proof['IndependentIntegrals'],
        'ReviewedNumFailStates':proof['ReviewedNumFailStates'],
        'MinimumStrictProofMarginN':min(x['StrictMarginN'] for x in proof['Failures']),
        'RunSHA256':sha(D/'Run.json'),'AcceptanceSHA256':sha(D/'Acceptance.json'),
        'StartedUTC':run['StartedUTC'],'FinishedUTC':run['FinishedUTC'],'ElapsedSeconds':run['ElapsedSeconds']}
(D/'Review.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps({k:v for k,v in result.items() if k not in ('Groups','NegativeControls','ExactHistoricalInputs')},ensure_ascii=False))
