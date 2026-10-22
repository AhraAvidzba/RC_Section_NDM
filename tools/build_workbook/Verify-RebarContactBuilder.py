"""Адресная приёмка CRebarContactBuilder и удаления заменённого модуля."""
import argparse
import hashlib
import importlib.util
import json
import re
import subprocess
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile,ZIP_DEFLATED

ROOT=Path(__file__).resolve().parents[2]
D=ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/RebarContactBuilder_2026-10-17'
RECIPES={'src/Geometry/CSectionRecipe'+x+'.cls' for x in ('Circle','Rectangles','PointedSection','HollowRectangle')}
NEW='src/Section/CRebarContactBuilder.cls';RETIRED='src/Geometry/modSectionRecipeRows.bas'
MATRICES=['modTestSectionRecipes.Dump'+x+'Recipe'+suffix for x,suffix in [('Circle','Matrix'),('Rectangles','Matrix'),('Rounded','Matrix'),('Hollow','Matrix'),('Hollow','EdgeMatrix')]]
SNAPSHOTS=MATRICES+['modTestAudit02Snapshots.TestDumpRecipeRuleContracts','modTestAudit02Snapshots.TestDumpRecipePathContracts','modTestAudit02Api.TestDumpAudit02CloseoutInputContract']
spec=importlib.util.spec_from_file_location('guide',Path(__file__).with_name('Verify-SectionRecipeGuide.py'))
guide=importlib.util.module_from_spec(spec);spec.loader.exec_module(guide)


def data(name):
    if (D/name).exists():return (D/name).read_bytes()
    with ZipFile(D/'Evidence.zip') as z:return z.read(name)


def text(name):return data(name).decode('utf-8-sig').replace('\r\n','\n')
def receipt(name):return json.loads(text(name))
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest().upper()
def write(name,value):(D/name).write_text(json.dumps(value,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')


def executable(body):
    return ''.join(p if p.startswith('"') else re.sub(r'\s+','',p.lower()) for p in re.findall(r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27]+',body) if not p.startswith("'"))


def source_scope():
    current={p.relative_to(ROOT).as_posix():p.read_text(encoding='utf-8-sig') for folder in ('src','tests','docs/geometry_refactor/examples') for p in (ROOT/folder).rglob('*') if p.suffix in ('.cls','.bas')}
    with ZipFile(D/'BeforeSources.zip') as z:
        old={name:z.read(name).decode('utf-8-sig').replace('\r\n','\n') for name in z.namelist()}
    assert current.keys()-old.keys()=={NEW}
    assert old.keys()-current.keys()=={RETIRED}
    changed={name for name in old.keys()&current.keys() if old[name]!=current[name]}
    assert changed==RECIPES|{'src/Section/CRebarRuleOptions.cls','src/Common/modSectionRecipeApiGuide.bas','tests/modTestAudit02Api.bas'},changed
    method=r'(?ms)^Public Sub ConfigureDependentContact\(.*?^End Sub'
    assert re.search(method,old['src/Section/CRebarRuleOptions.cls'])[0]==re.search(method,current['src/Section/CRebarRuleOptions.cls'])[0]
    fields=re.findall(r'^Public (\w+) As ',current['src/Section/CRebarRuleOptions.cls'],re.M)
    clone=re.search(r'(?ms)^Public Function Clone\(.*?^End Function',current['src/Section/CRebarRuleOptions.cls'])[0]
    assert set(re.findall(r'\bresult\.(\w+) = ',clone))==set(fields)
    source=current[NEW]
    assert re.findall(r'(?m)^Public (?:Sub|Function) (\w+)\(',source)==['Initialize','AppendRow']
    assert source.count('.ConfigureDependentContact ')==1 and '/ 2' not in source
    assert not re.search(r'\b(?:Application|ThisWorkbook|Worksheet|CSectionDefinitionBuilder|CSectionGeometryQuery|CRebarRuleExecutor)\b',executable(source),re.I)
    for name in RECIPES:
        body=current[name]
        # Рецепты передают точные размеры; шаг подписи остаётся декларативным параметром.
        assert not re.search(r'\b(?:Round|Format\$?|FormatNumber|Fix|Int|CInt|CLng|WorksheetFunction)\s*\(',executable(body),re.I),name
        add=re.search(r'(?ms)^Private Sub AddReinforcement\(.*?^End Sub',body)[0]
        assert 'AppendRow' not in add and 'ConfigureDependentContact' not in add and 'CRebarContactBuilder' not in add
        assert 'ExtraOptions(rowNumber)' in add and 'facade.RebarDependent' in add
        assert '.Main"' in add and 'SkippedDiameters' not in body
        assert body.count('New CRebarContactBuilder')==body.count('.AppendRow(')==1
    for name,body in current.items():assert 'RecipeContactOptions' not in body and 'modSectionRecipeRows' not in body,name
    assert receipt('CommentAudit.json')['status']=='PASS'
    api=json.loads((ROOT/'docs/geometry_refactor/SECTION_RECIPE_API_MANIFEST.json').read_text(encoding='utf-8-sig'))
    baseline=json.loads(subprocess.check_output(['git','show',receipt('Before.json')['Head']+':docs/geometry_refactor/SECTION_RECIPE_API_MANIFEST.json'],cwd=ROOT))
    a,b=[{m['id']:{k:v for k,v in m.items() if k!='line'} for m in item['members']} for item in (baseline,api)]
    assert b.keys()-a.keys()=={'CRebarContactBuilder.Initialize','CRebarContactBuilder.AppendRow','CRebarRuleOptions.Clone'}
    assert a.keys()-b.keys()=={'modSectionRecipeRows.RecipeContactOptions'}
    assert all(a[name]==b[name] for name in a.keys()&b.keys())
    return {'Added':[NEW],'Removed':[RETIRED],'Changed':sorted(changed),'ExistingContactFormulaUnchanged':True,'ClonedFields':len(fields),'OnlyOneNewClass':True}


def run(stage,expected=None):
    record=receipt(stage+'/Run.json')
    assert record['Passed'] and record['ProductionUnchanged'],stage
    if expected is not None:assert [r['Macro'] for r in record['Runs']]==expected
    count=0
    for item in record['Runs']:
        body=text(stage+'/'+item['Macro']+'.txt')
        assert item['Passed'] and body.strip() and not re.search(r'(?m)^FAIL[: ]|failed=[1-9]|RUNTIME ERROR:|^ERROR:',body),stage
        totals=re.findall(r'(?m)^TOTAL[^\n]*passed=(\d+); failed=0\b',body)
        if not item['Snapshot']:assert totals;count+=int(totals[-1])
    return {'Stage':stage,'Assertions':count,'SourceSHA256':record['SourceSHA256']}


def components(body):
    return {m[1]:m[3] for m in re.finditer(r'(?ms)^COMPONENT: ([^\n]+)\nTYPE: ([^\n]+)\nLINES: [^\n]+\n=+\n(.*?)(?=\n=+\nCOMPONENT: |\Z)',body)}


def tests():
    runs=[run('BaselineMatrices',SNAPSHOTS),run('CandidateMatrices',SNAPSHOTS),run('CandidateUnit',['modTestAudit02Api.RunRebarContactBuilderTests']),run('Directed'),
          run('CompiledSmoke',['modTestAudit02Api.RunRebarContactBuilderTests','modTestSectionRecipes.RunWorkbookRecipeInputTests','modTestSectionRecipes.RunCircleBindingWorkbookTests'])]
    comparisons=[]
    for macro in SNAPSHOTS:
        name=macro+'.txt';a,b=data('BaselineMatrices/'+name),data('CandidateMatrices/'+name)
        assert a==b,macro
        comparisons.append({'Macro':macro,'Bytes':len(b),'SHA256':hashlib.sha256(b).hexdigest().upper()})
    unit=text('CandidateUnit/modTestAudit02Api.RunRebarContactBuilderTests.txt')
    assert 'TOTAL REBAR CONTACT BUILDER: passed=97; failed=0' in unit
    assert all('PASS: contactBuilder.'+name in unit for name in ('example.first','example.second','example.third','example.nominal','snapshot.setbacks','snapshot.endMode','snapshot.names','snapshot.source','rows.independent','rows.firstRetained','uninitialized','reject.location','reject.binding','reject.zero','reject.negative','reject.reinitialize','overflow.recover.0','overflow.recover.1'))
    assert text('CandidateMatrices/modTestAudit02Snapshots.TestDumpRecipeRuleContracts.txt').count('CASE|')==60
    measurements=text('Directed/modTestSectionRecipes.RunRecipeRuleTests.txt')
    for name in ('round.positive','round.negative','round.exact','coordinates.unchanged'):
        assert 'PASS: measurements.'+name+';' in measurements
    a,b=components(text('CandidateNativeVBA.txt')),components(text('CompiledNativeVBA.txt'))
    assert a.keys()==b.keys() and len(b)==158 and 'CRebarContactBuilder' in b and 'modSectionRecipeRows' not in b
    assert all(executable(a[name])==executable(b[name]) for name in a),'Compiled working code changed after tests.'
    before=components(subprocess.check_output(['git','show',receipt('Before.json')['Head']+':workbook/output/VBA_All_Code.txt'],cwd=ROOT).decode('utf-8-sig').replace('\r\n','\n'))
    assert b.keys()-before.keys()=={'CRebarContactBuilder'} and before.keys()-b.keys()=={'modSectionRecipeRows'}
    changed={name for name in before.keys()&b.keys() if executable(before[name])!=executable(b[name])}
    assert changed=={'CSectionRecipe'+x for x in ('Circle','Rectangles','PointedSection','HollowRectangle')}|{'CRebarRuleOptions','modTestAudit02Api'},changed
    assert receipt('CompiledSourceEquality.json')['Passed']
    return {'Runs':runs,'DirectedAssertions':runs[2]['Assertions']+runs[3]['Assertions'],'SmokeAssertions':runs[4]['Assertions'],'Comparisons':comparisons,'RecipeCases':60,
            'Components':158,'LegacyModuleAbsent':True,'OtherExecutableComponentsUnchanged':151,'RecipesDoNotRoundDimensions':True,'DisplayRoundingPreservesCoordinates':True}


def package():
    scope=source_scope();tested=tests()
    before=guide.read(D/'Baseline.xlsm')
    old=before['sheets']['Справка']['cells']['A574']['value']
    new=next(line for line in (ROOT/'docs/geometry_refactor/ADDING_SECTION_TYPE.md').read_text(encoding='utf-8-sig').splitlines() if line.startswith('4. '))
    assert 'RecipeContactOptions' in old and 'CRebarContactBuilder' in new and len(new)<len(old)
    with ZipFile(D/'Baseline.xlsm') as original,ZipFile(D/'Compiled.xlsm') as compiled,ZipFile(D/'Accepted.xlsm','w',ZIP_DEFLATED) as result:
        vba=compiled.read('xl/vbaProject.bin')
        for info in original.infolist():
            body=original.read(info.filename)
            if info.filename=='xl/vbaProject.bin':body=vba
            if info.filename=='xl/sharedStrings.xml':
                x,y=escape(old).encode('utf-8'),escape(new).encode('utf-8');assert body.count(x)==1;body=body.replace(x,y)
            result.writestr(info,body)
    after=guide.read(D/'Accepted.xlsm');cells=0
    assert before['names']==after['names'] and before['drawings']==after['drawings']
    for sheet,a in before['sheets'].items():
        b=after['sheets'][sheet];assert a['rows']==b['rows'] and a['cells'].keys()==b['cells'].keys()
        for address,value in a['cells'].items():
            expected=dict(value)
            if sheet=='Справка' and address=='A574':expected['value']=new
            assert expected==b['cells'][address],(sheet,address)
            cells+=1
        assert guide.canonical(a['document'])==guide.canonical(b['document'])
    with ZipFile(D/'Baseline.xlsm') as a,ZipFile(D/'Accepted.xlsm') as b:
        assert a.namelist()==b.namelist()
        assert sorted(name for name in a.namelist() if a.read(name)!=b.read(name))==['xl/sharedStrings.xml','xl/vbaProject.bin']
    record={'Passed':True,'SourceScope':scope,'Tests':tested,'AcceptedSHA256':sha(D/'Accepted.xlsm'),'CompiledSHA256':sha(D/'Compiled.xlsm'),
            'CompiledVbaSHA256':hashlib.sha256(vba).hexdigest().upper(),'PreservedCells':cells,'PreservedNames':len(before['names']),'HelpChange':{'Cell':'Справка!A574','Before':old,'After':new}}
    write('Package.json',record);return record


def accept(published):
    scope=source_scope();tested=tests();package_data=receipt('Package.json')
    assert package_data['Passed'] and sha(D/'Accepted.xlsm')==package_data['AcceptedSHA256']
    assert receipt('HelpReview.json')['Passed']
    assert sha(ROOT/'workbook/output/RC_Section_NDM_execution_report.txt')==receipt('Before.json')['UserReportSHA256']
    if published:
        assert sha(ROOT/'workbook/output/RC_Section_NDM.xlsm')==package_data['AcceptedSHA256']
        assert receipt('PublishedSourceEquality.json')['Passed']
        assert (ROOT/'workbook/output/VBA_All_Code.txt').read_bytes()==data('PublishedNativeVBA.txt')
        body=text('PublishedNativeVBA.txt');assert components(body)==components(text('CompiledNativeVBA.txt'))
        assert not re.search(r'NDM\.RECIPE\.GUIDE|[0-9a-fA-F]{64}|RecipeContactOptions|modSectionRecipeRows',body)
        with ZipFile(ROOT/'workbook/output/RC_Section_NDM.xlsm') as z:
            assert all('RecipeContactOptions' not in z.read(name).decode('utf-8-sig') and 'modSectionRecipeRows' not in z.read(name).decode('utf-8-sig') for name in z.namelist() if name.endswith('.xml'))
    record={'Passed':True,'PublishedVerified':published,'SourceScope':scope,'Tests':tested,'PreservedCells':package_data['PreservedCells'],'PreservedNames':package_data['PreservedNames'],'UserReportUnchanged':True,'FullOnOffRun':False}
    write('Acceptance.json',record);return record


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--package',action='store_true');parser.add_argument('--published',action='store_true');args=parser.parse_args()
    result=package() if args.package else accept(args.published)
    print(json.dumps({k:v for k,v in result.items() if k not in ('SourceScope','Tests','HelpChange')},ensure_ascii=False))
