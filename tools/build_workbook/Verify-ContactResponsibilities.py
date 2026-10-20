"""Приёмка переноса подготовки контактных рядов в AddReinforcement."""
import argparse
import hashlib
import importlib.util
import json
import re
import subprocess
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile, ZIP_DEFLATED

ROOT=Path(__file__).resolve().parents[2]
D=ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/ContactResponsibilities_2026-10-17'
SHAPES=('Circle','Rectangles','PointedSection','HollowRectangle')
RECIPES={'src/Geometry/CSectionRecipe'+x+'.cls' for x in SHAPES}
SNAPSHOTS=['modTestSectionRecipes.Dump'+x+'Recipe'+suffix for x,suffix in [('Circle','Matrix'),('Rectangles','Matrix'),('Rounded','Matrix'),('Hollow','Matrix'),('Hollow','EdgeMatrix')]]
SNAPSHOTS+=['modTestAudit02Snapshots.DumpRecipeRuleContracts','modTestAudit02Snapshots.DumpRecipePathContracts','modTestAudit02Api.DumpAudit02CloseoutInputContract']
spec=importlib.util.spec_from_file_location('previous',Path(__file__).with_name('Verify-RebarContactBuilder.py'))
previous=importlib.util.module_from_spec(spec);spec.loader.exec_module(previous)
guide=previous.guide


def data(name):
    p=D/name
    if p.exists():return p.read_bytes()
    with ZipFile(D/'Evidence.zip') as z:return z.read(name)
def text(name):return data(name).decode('utf-8-sig').replace('\r\n','\n')
def receipt(name):return json.loads(text(name))
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest().upper()
def write(name,value):(D/name).write_text(json.dumps(value,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
def old(path):return subprocess.check_output(['git','show',receipt('Before.json')['Head']+':'+path],cwd=ROOT)


def source_scope():
    changed=set(subprocess.check_output(['git','diff',receipt('Before.json')['Head'],'--name-only','--','src','tests','docs/geometry_refactor/examples'],cwd=ROOT).decode().splitlines())
    assert changed==RECIPES|{'src/Common/modSectionRecipeApiGuide.bas'},changed
    for name in RECIPES:
        body=(ROOT/name).read_text(encoding='utf-8-sig')
        methods=dict(re.findall(r'(?ms)^Private Sub (\w+)\((.*?^End Sub)',body))
        for method,part in methods.items():
            if method.startswith('Read'):
                assert not re.search(r'\b(?:CRebarContactBuilder|CRebarRuleOptions|AppendRow|ConfigureDependentContact|facade)\b',part),name
        add=methods['AddReinforcement']
        assert add.count('New CRebarContactBuilder')==add.count('contact.Initialize ')==add.count('contact.AppendRow(')==1
        assert body.count('New CRebarContactBuilder')==body.count('contact.AppendRow(')==1
        assert 'ExtraOptions' not in body and 'ExtraOptions' not in add
        assert 'table.' not in add and 'context.' not in add
        assert 'Set defaults = New CRebarRuleOptions' in add and '.Clone' not in add
        assert 'New CRebarContactBuilder' in add and 'facade.RebarDependent' in add and '.Main"' in add
        assert not re.search(r'\b(?:Round|Format\$?|FormatNumber|Fix|Int|CInt|CLng)\s*\(',previous.executable(body),re.I)
        for udt in re.findall(r'(?ms)^Private Type .*?^End Type',body):
            assert not re.search(r'\bAs C\w+',udt),name
    for name in ('CRebarContactBuilder','CRebarRuleOptions'):
        path='src/Section/'+name+'.cls'
        assert (ROOT/path).read_text(encoding='utf-8-sig').replace('\r\n','\n')==old(path).decode('utf-8-sig').replace('\r\n','\n')
    assert receipt('CommentAudit.json')['status']=='PASS'
    return {'Recipes':len(RECIPES),'ReadMethodsCreateNoOptions':True,'BuilderOnlyInAddReinforcement':True,'InputStructuresHaveNoObjects':True,'ContactClassAndFormulaUnchanged':True,'RecipesDoNotRoundDimensions':True}


def tests():
    for stage in ('BaselineMatrices','CandidateMatrices','Directed'):
        run=receipt(stage+'/Run.json')
        assert run['Passed'] and run['ProductionUnchanged'],stage
        if stage!='Directed':assert [x['Macro'] for x in run['Runs']]==SNAPSHOTS
        assert run['SourceSHA256']==(receipt('Before.json')['WorkbookSHA256'] if stage=='BaselineMatrices' else receipt('CodeUpdate.json')['WorkbookSHA256'])
        for item in run['Runs']:
            body=text(stage+'/'+item['Macro']+'.txt')
            assert item['Passed'] and body.strip() and not re.search(r'(?m)^FAIL[: ]|failed=[1-9]|RUNTIME ERROR:|^ERROR:',body),item['Macro']
    comparisons=[]
    for macro in SNAPSHOTS:
        name=macro+'.txt';a,b=data('BaselineMatrices/'+name),data('CandidateMatrices/'+name)
        assert a==b,macro
        comparisons.append({'Macro':macro,'Bytes':len(b),'SHA256':hashlib.sha256(b).hexdigest().upper()})
    assert text('CandidateMatrices/modTestAudit02Snapshots.DumpRecipeRuleContracts.txt').count('CASE|')==60
    assertions=0
    for item in receipt('Directed/Run.json')['Runs']:
        totals=re.findall(r'(?m)^TOTAL[^\n]*passed=(\d+); failed=0\b',text('Directed/'+item['Macro']+'.txt'))
        assert totals;assertions+=int(totals[-1])
    assert 'TOTAL REBAR CONTACT BUILDER: passed=97; failed=0' in text('Directed/modTestAudit02Api.RunRebarContactBuilderTests.txt')
    before=previous.components(old('workbook/output/VBA_All_Code.txt').decode('utf-8-sig').replace('\r\n','\n'))
    after=previous.components(text('CandidateNativeVBA.txt'))
    assert len(before)==len(after)==158 and before.keys()==after.keys()
    changed={name for name in before if previous.executable(before[name])!=previous.executable(after[name])}
    assert changed=={'CSectionRecipe'+x for x in SHAPES},changed
    assert receipt('CodeUpdate.json')['Compiled'] and receipt('AcceptedSourceEquality.json')['Passed']
    assert 'modSectionRecipeRows' not in after and 'CRebarContactBuilder' in after
    return {'Assertions':assertions,'Comparisons':comparisons,'RecipeCases':60,'Components':158,'OtherExecutableComponentsUnchanged':154,'NoFullOnOff':True}


def package():
    scope=source_scope();checked=tests()
    assert sha(D/'Candidate.xlsm')==receipt('CodeUpdate.json')['WorkbookSHA256']
    before=guide.read(D/'Baseline.xlsm')
    old_text=before['sheets']['Справка']['cells']['A574']['value']
    new_text=next(x for x in (ROOT/'docs/geometry_refactor/ADDING_SECTION_TYPE.md').read_text(encoding='utf-8-sig').splitlines() if x.startswith('4. '))
    with ZipFile(D/'Baseline.xlsm') as original,ZipFile(D/'Candidate.xlsm') as candidate,ZipFile(D/'Accepted.xlsm','w',ZIP_DEFLATED) as result:
        for item in original.infolist():
            body=original.read(item.filename)
            if item.filename=='xl/vbaProject.bin':body=candidate.read(item.filename)
            if item.filename=='xl/sharedStrings.xml':
                a,b=escape(old_text).encode('utf-8'),escape(new_text).encode('utf-8');assert body.count(a)==1;body=body.replace(a,b)
            result.writestr(item,body)
    after=guide.read(D/'Accepted.xlsm');cells=0
    assert before['names']==after['names'] and before['drawings']==after['drawings']
    for sheet,a in before['sheets'].items():
        b=after['sheets'][sheet];assert a['rows']==b['rows'] and a['cells'].keys()==b['cells'].keys()
        for address,value in a['cells'].items():
            expected=dict(value)
            if sheet=='Справка' and address=='A574':expected['value']=new_text
            assert expected==b['cells'][address],(sheet,address)
            cells+=1
        assert guide.canonical(a['document'])==guide.canonical(b['document'])
    with ZipFile(D/'Baseline.xlsm') as a,ZipFile(D/'Accepted.xlsm') as b:
        assert a.namelist()==b.namelist()
        assert sorted(n for n in a.namelist() if a.read(n)!=b.read(n))==['xl/sharedStrings.xml','xl/vbaProject.bin']
    record={'Passed':True,'Scope':scope,'Tests':checked,'AcceptedSHA256':sha(D/'Accepted.xlsm'),'PreservedCells':cells,'PreservedNames':len(before['names']),'HelpBefore':old_text,'HelpAfter':new_text}
    write('Package.json',record);return record


def accept(published):
    scope=source_scope();checked=tests();package_data=receipt('Package.json')
    assert package_data['Passed'] and sha(D/'Accepted.xlsm')==package_data['AcceptedSHA256']
    assert receipt('HelpReview.json')['Passed']
    assert sha(ROOT/'workbook/output/RC_Section_NDM_execution_report.txt')==receipt('Before.json')['UserReportSHA256']
    if published:
        assert sha(ROOT/'workbook/output/RC_Section_NDM.xlsm')==package_data['AcceptedSHA256']
        assert receipt('PublishedSourceEquality.json')['Passed']
        assert (ROOT/'workbook/output/VBA_All_Code.txt').read_bytes()==data('PublishedNativeVBA.txt')
        assert previous.components(text('CandidateNativeVBA.txt'))==previous.components(text('PublishedNativeVBA.txt'))
    result={'Passed':True,'PublishedVerified':published,'Scope':scope,'Tests':checked,'PreservedCells':package_data['PreservedCells'],'PreservedNames':package_data['PreservedNames'],'UserReportUnchanged':True}
    write('Acceptance.json',result);return result


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--package',action='store_true');parser.add_argument('--published',action='store_true');args=parser.parse_args()
    result=package() if args.package else accept(args.published)
    print(json.dumps({k:v for k,v in result.items() if k not in ('Scope','Tests','HelpBefore','HelpAfter')},ensure_ascii=False))
