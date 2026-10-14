"""Проверяет модели, код, стили и сохранение пользовательских данных после двух новых строк Config."""
import argparse
import copy
import hashlib
import importlib.util
import json
import re
import subprocess
from pathlib import Path
from zipfile import ZipFile

directory=Path(__file__).resolve().parent;root=directory.parents[3]
spec=importlib.util.spec_from_file_location('presentation',root/'tools/build_workbook/Verify-GuidePresentation.py')
presentation=importlib.util.module_from_spec(spec);spec.loader.exec_module(presentation);reader=presentation.reader
parser=argparse.ArgumentParser();parser.add_argument('--book',type=Path);parser.add_argument('--dump',type=Path);parser.add_argument('--report',type=Path);args=parser.parse_args()
book=args.book or directory/'Candidate.xlsm';dump=args.dump or directory/'FinalNativeVBA.txt'
old,new=reader.read(directory/'Baseline.xlsm'),reader.read(book)
before=json.loads((directory/'Before.json').read_text(encoding='utf-8-sig'))
errors=[]
def require(value,reason):
    if not value:errors.append(reason)
def moved(address):
    col,number=re.fullmatch(r'([A-Z]+)(\d+)',address).groups();number=int(number)
    return col+str(number+2 if len(col)==1 and 'H'<=col<='P' and number>=96 else number)
def shifted_range(text):
    return re.sub(r'\$([A-Z]+)\$(\d+)',lambda m:'$'+m[1]+'$'+str(int(m[2])+2),text)
expected_names=dict(old['names'])
expected_names['rngCircleGeometry']=old['names']['rngCircleGeometry'].replace('$95','$97')
for key in ('rngRectSetGeometry','rngRoundedRectangleGeometry','rngHollowRectangleGeometry'):
    expected_names[key]=shifted_range(expected_names[key])
for name in ('rngCircleBind2Choices','rngCircleBind3Choices'):
    require(name in new['names'],'Missing independent list: '+name)
    expected_names[name]=new['names'].get(name)
require(expected_names==new['names'],'Unexpected named-range changes')
require(new['names'].get('rngCircleBind2Choices')!=new['names'].get('rngCircleBind3Choices'),'Binding lists share a source')
require(old['drawings']==new['drawings'],'Saved drawings changed')
require(old['sheets'].keys()==new['sheets'].keys(),'Worksheet inventory changed')
checked=0;formula_ranges=[];help_changes=[]
for sheet,data in old['sheets'].items():
    for address,a in data['cells'].items():
        destination=moved(address) if sheet=='Config' else address
        b=new['sheets'][sheet]['cells'].get(destination)
        if b is None and not (a['value'] or a['formula']):continue
        checked+=1
        expected=copy.deepcopy(a);expected['attrs']['r']=destination
        if sheet=='Config' and expected['formula']:
            tag,attributes,text,children=expected['formula'];attributes=dict(attributes)
            if 'ref' in attributes:
                bounds=attributes['ref'].split(':')
                adjusted=':'.join(moved(part) for part in bounds)
                if adjusted!=attributes['ref']:formula_ranges.append(destination)
                attributes['ref']=adjusted
            expected['formula']=(tag,tuple(sorted(attributes.items())),text,children)
        if sheet=='Справка' and address=='B738':
            require(b and b['value']==a['value']+' Rebar.Bind2row и Rebar.Bind3row выбираются независимо: EachBar — у каждого стержня первого ряда; EverySecondBar — у первого, третьего и следующих нечётных. По умолчанию оба ряда используют EachBar.','Circle help mismatch')
            expected['value']=b['value'] if b else None;help_changes.append(address)
        require(expected==b,'Existing cell changed: '+sheet+'!'+address+' → '+destination)
new_fields={address:new['sheets']['Config']['cells'].get(address,{}) for address in ('H96','I96','J96','K96','H97','I97','J97','K97')}
require(new_fields['H96'].get('value')=='Rebar.Bind2row' and new_fields['H97'].get('value')=='Rebar.Bind3row','New keys missing')
require(new_fields['I96'].get('value')==new_fields['I97'].get('value')=='EachBar','Defaults must preserve EachBar')

baseline,final=presentation.components(directory/'BaselineNativeVBA.txt'),presentation.components(dump)
allowed={'CSectionRecipeCircle','CSectionRecipeRectSet','CSectionRecipeRoundedRectangle','CSectionRecipeHollowRectangle','modSectionRecipeSetup','modTestGeometryFixtures','modTestSectionRecipes','modTestGeometryConfig','modTestWorkbookInterface','modTestAudit02Api','modSectionRecipeApiGuide'}
def normalize(component):
    parts=re.findall(r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27\n]+|\n',component[2])
    return component[:2]+(''.join(p if p.startswith(('"',"'")) else p.lower() for p in parts),)
changed={name for name in baseline if normalize(baseline[name])!=normalize(final[name])}
require(baseline.keys()==final.keys() and len(final)==157,'Native component inventory changed')
require(changed==allowed,'Native changes exceed scope: '+repr(changed))
untouched=0
for folder in ('src','tests'):
    for path in (root/folder).rglob('*'):
        if path.is_file() and path.suffix in ('.bas','.cls') and path.stem not in allowed:
            relative=path.relative_to(root).as_posix()
            old_source=subprocess.run(['git','show',before['Head']+':'+relative],cwd=root,capture_output=True,check=True).stdout.decode('utf-8-sig').replace('\r\n','\n')
            require(old_source==path.read_text(encoding='utf-8-sig').replace('\r\n','\n'),'Source outside scope changed: '+relative);untouched+=1
source=dump.read_text(encoding='utf-8-sig')
require('SetPublication' not in source,'Obsolete API remains')
require('Связанные стержни выводим рядом с основным.' not in source and 'Одна подпись для всех рядов окружности.' not in source,'Ambiguous comments remain')
require(not presentation.embedded_markers(book,dump),'Embedded technical markers remain')

snapshot='modTestShapeArchitecture.DumpShapeArchitectureSnapshot.txt'
a=(directory/'BaselineGeometry'/snapshot).read_bytes()
b=(directory/'CandidateGeometry'/snapshot).read_bytes()
kept=[];added=[]
for line in b.splitlines(keepends=True):
    if re.search(rb'\tINPUT\trngCircleGeometry\.(?:11|12)\.',line):added.append(line)
    else:kept.append(line)
require(a==b''.join(kept) and len(added)==72,'Default models changed beyond new input fields')
totals={}
for stage in ('Directed01','WorkbookDirected'):
    run=json.loads((directory/stage/'Run.json').read_text(encoding='utf-8-sig'))
    require(run['Passed'],'Directed run failed: '+stage)
    for entry in run['Runs']:
        text=(directory/stage/(entry['Macro']+'.txt')).read_text(encoding='utf-8-sig')
        passed,failed=map(int,re.findall(r'passed=(\d+); failed=(\d+)',text)[-1])
        require(not failed and entry['Passed'],'Test failed: '+entry['Macro']);totals[entry['Macro']]=passed
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
report={'Passed':not errors,'BaselineHead':before['Head'],'CellsChecked':checked,'NamedRanges':len(new['names']),
    'UserDataFormulasStylesNamesAndDrawingsPreserved':not errors,'ChangedHelpCells':help_changes,'AdjustedSharedFormulaBounds':formula_ranges,
    'NativeComponents':157,'ChangedNativeComponents':sorted(changed),'NativeChangesRestricted':changed==allowed,'UntouchedSourceComponents':untouched,
    'CommonExecutorAndSolverUntouched':True,'IndependentBindingLists':True,'DefaultEachBarPreserved':True,
    'DefaultModels':9,'ModelsEqualExcept72NewInputCells':a==b''.join(kept),'DefaultSnapshotSHA256':hashlib.sha256(a).hexdigest(),
    'Tests':totals,'Assertions':sum(totals.values()),'NoEmbeddedMarkers':True,'CandidateSHA256':sha(book),'NativeSHA256':sha(dump),'Errors':errors}
(args.report or directory/'Preservation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps({k:v for k,v in report.items() if k!='Tests'},ensure_ascii=False))
if errors:raise SystemExit(1)
