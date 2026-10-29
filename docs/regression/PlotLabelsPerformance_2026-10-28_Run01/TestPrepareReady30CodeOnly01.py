"""ДЛЯ ТЕСТОВ: готовит точное обновление VBA тяжёлой книги только после полного PASS и выхода всех own."""
from pathlib import Path
import argparse,hashlib,json,zipfile
parser=argparse.ArgumentParser()
parser.add_argument('--launch-directory',required=True)
args=parser.parse_args()
e=Path(__file__).resolve().parent;root=e.parents[2]
launch=(root/args.launch_directory).resolve()
assert launch.is_relative_to(root/'docs/regression')
def read(f):return json.loads(f.read_text('utf-8-sig'))
def sha(f):return hashlib.sha256(f.read_bytes()).hexdigest().upper()
request=read(launch/'LaunchRequest.json')
core=root/request['EvidenceDirectory']
complete=read(core/'CompleteCurrentAcceptance.json')
assert complete['Status']=='PASS' and complete['SourceSHA256']==request['SourceSHA256']
assert complete['CoreSuites']==182 and complete['FocusedSingleCases']==122
own=read(launch/'AllOwnClosuresReviewed.json')
assert own['AllOwnIdentitiesTerminal'] and own['PrimaryLogsReviewed'] and own['SourceSHA256']==request['SourceSHA256']
assert own['CoreMacrosReviewed']==182 and own['FocusedSingleCasesReviewed']==122 and own['FocusedRecordCount']==26
main=root/'workbook/output/RC_Section_NDM.xlsm'
assert sha(main)==request['SourceSHA256']
plan=read(e/'Ready30FinalPlan.json');source=Path(plan['HistoricalThirtySource'])
inspection=read(e/'Ready30SourceInspection01.json')
assert inspection['Status']=='PASS_PASSIVE_READY_SOURCE_INSPECTION'
structure=read(e/'Ready30StructureComparison01.json')
assert structure['Status']=='PASS_PASSIVE_STRUCTURE_CONTRACT'
assert structure['CurrentSHA256']==request['SourceSHA256'], 'Repeat passive structure/list comparison on the exact accepted main first'
assert structure['SourceSHA256']==inspection['SourceSHA256']
assert not structure['DifferentFixedValidations'] and not structure['DifferentAnchors']
assert inspection['Counts']=={'Concrete':1232,'Rebar':8} and len(inspection['LoadIDs'])==30
assert sha(source)==inspection['SourceSHA256']
assert sha(root/plan['PreserveUserOtherCopy'])==plan['PreserveUserOtherCopySHA256']
p=Path('C:/Users/avidzba/.codex/visualizations/2026/10/12/01a12ece-8192-7141-89b3-f83f1176184c/PlotLabelsPerformance_2026-10-28_Run01/Ready30CodeOnly01')
p.mkdir(exist_ok=False)
out=p/'RC_Section_NDM.xlsm'
with zipfile.ZipFile(main) as z:vba=z.read('xl/vbaProject.bin')
with zipfile.ZipFile(source) as old,zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as new:
    names=old.namelist();assert len(names)==len(set(names))
    assert not any('signature' in n.lower() for n in names), 'Signed source needs a separate reviewed packaging path'
    for item in old.infolist():new.writestr(item,vba if item.filename=='xl/vbaProject.bin' else old.read(item.filename))
with zipfile.ZipFile(source) as old,zipfile.ZipFile(out) as new:
    assert new.testzip() is None and old.namelist()==new.namelist()
    preserved=[]
    for name in old.namelist():
        if name=='xl/vbaProject.bin':assert new.read(name)==vba;continue
        assert old.read(name)==new.read(name)
        preserved.append({'Part':name,'Bytes':len(old.read(name)),'SHA256':hashlib.sha256(old.read(name)).hexdigest().upper()})
proof={'Status':'PREPARED_NOT_PUBLISHED_NATIVE_REVIEW_REQUIRED','AcceptedMainSHA256':sha(main),'AcceptedCommit':request['SourceCommit'],
       'Source':str(source),'SourceSHA256':sha(source),'Ready':str(out),'ReadySHA256':sha(out),'VBAProjectSHA256':hashlib.sha256(vba).hexdigest().upper(),
       'OnlyVBAChanged':True,'DataAndLoadsRecalculated':False,'PreservedParts':preserved,'PreservedCount':len(preserved),
       'UserOtherCopyUnchanged':True,'FullSourcePassVerified':True,'CurrentFullOwnClosed':True}
(e/'Ready30CodePreparation01.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({k:v for k,v in proof.items() if k!='PreservedParts'},ensure_ascii=False,indent=2))
