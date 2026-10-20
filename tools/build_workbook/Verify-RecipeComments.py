"""Проверяет исправления пояснений рецептов без изменения исполняемых команд."""
import argparse
import hashlib
import importlib.util
import json
import math
import re
import struct
import subprocess
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED

ROOT=Path(__file__).resolve().parents[2]
D=ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/RecipeComments_2026-10-16'
PREVIOUS=ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/RecipeResponsibilities_2026-10-16'
RECIPES=['src/Geometry/CSectionRecipe'+name+'.cls' for name in ('Circle','Rectangles','PointedSection','HollowRectangle')]
EXAMPLES=['docs/geometry_refactor/examples/CSectionRecipe'+name+'.cls' for name in ('MySection','RoundedTrapezoid')]
RELATED=['src/Section/ISectionRecipe.cls','src/Section/CSectionRecipeContext.cls','src/Geometry/modSectionRecipeRows.bas','src/Section/CRebarRuleOptions.cls']
CHANGED=set(RECIPES+['src/Section/CRebarRuleOptions.cls','src/Common/modSectionRecipeApiGuide.bas',EXAMPLES[1]])
spec=importlib.util.spec_from_file_location('comments',Path(__file__).with_name('Audit-ClassComments.py'))
comments=importlib.util.module_from_spec(spec);spec.loader.exec_module(comments)


def data(name):
    if (D/name).exists():return (D/name).read_bytes()
    with ZipFile(D/'Evidence.zip') as z:return z.read(name)


def text(name):return data(name).decode('utf-8-sig').replace('\r\n','\n')
def receipt(name):return json.loads(text(name))
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest().upper()
def write(name,value):(D/name).write_text(json.dumps(value,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')


def format_cases(source,name):
    if name not in ('CSectionRecipePointedSection','CSectionRecipeHollowRectangle'):return source,0
    pattern=r'(?m)^( +)Case ("B.Top"|"B.Bottom"|3|4): (Set AxisPath = facade\.Chain[^\n]*)$'
    return re.subn(pattern,lambda m:m[1]+'Case '+m[2]+'\n'+m[1]+'    '+m[3],source)


def source_scope():
    baseline=receipt('Before.json')['Head']
    changed=[];formatted=[];total=0
    paths=[p for folder in ('src','tests','docs/geometry_refactor/examples') for p in (ROOT/folder).rglob('*') if p.suffix in ('.cls','.bas')]
    for p in paths:
        name=p.relative_to(ROOT).as_posix()
        old=subprocess.check_output(['git','show',baseline+':'+name],cwd=ROOT).decode('utf-8-sig').replace('\r\n','\n')
        current=p.read_text(encoding='utf-8-sig')
        if old!=current:changed.append(name)
        prepared,n=format_cases(old,p.stem)
        assert comments.executable_lines(prepared)==comments.executable_lines(current),name
        if n:formatted.append({'File':name,'CaseStatements':n})
        total+=1
    assert set(changed)==CHANGED,changed
    assert sum(x['CaseStatements'] for x in formatted)==4
    reviewed=[]
    for name in RECIPES+EXAMPLES+RELATED:
        source=(ROOT/name).read_text(encoding='utf-8-sig')
        reviewed.append({'File':name,'Lines':len(source.splitlines()),'CommentLines':sum("'" in line for line in source.splitlines())})
    assert receipt('CommentAudit.json')['status']=='PASS'
    result={'Passed':True,'ComparedComponents':total,'ChangedFiles':sorted(changed),'OnlyCommentsAndCaseLayoutChanged':True,
            'CaseLayout':formatted,'ReviewedFiles':reviewed,'ReviewBasis':['CGeometryPath.BetweenAnchors','CSectionGeometryQuery.ConstrainLine',
            'CRebarRuleExecutor.StationsFor','CRebarRuleOptions.ConfigureDependentContact','CSectionAnnotationBuilder.Build']}
    write('SourceScope.json',result)
    return result


def directions():
    # Снимки сняты из той самой книги, которая является исходной для этой правки.
    previous=json.loads((PREVIOUS/'Package.json').read_text(encoding='utf-8-sig'))
    assert previous['AcceptedSHA256']==receipt('Before.json')['WorkbookSHA256']
    with ZipFile(PREVIOUS/'Evidence.zip') as z:
        lines=z.read('CompiledContracts/modTestAudit02Snapshots.DumpRecipePathContracts.txt').decode('utf-8-sig').splitlines()
    case=None;path=None;segments=[];checked=[]
    def flush():
        if not path or not segments:return
        first,last=segments[0],segments[-1]
        family=case.split('|')[1];direction=None
        if family.startswith('Rounded') and path in ('B.Top','B.Bottom','H.Left','H.Right'):
            direction='Right' if path.startswith('B.') else 'Down'
        elif family.startswith('Hollow') and path.removeprefix('Opening.') in ('B.Top','B.Bottom','H.Left','H.Right'):
            direction='Right' if '.B.' in '.'+path else 'Up'
        elif family in ('Rectangle','LSection','TwoRectangles') and re.fullmatch(r'[BH][12]\.as_[12]',path):
            direction=('Up' if path.endswith('_1') else 'Down') if path.startswith('H') else ('Right' if path.endswith('_1') else 'Left')
        elif family=='Circle' and path=='Circle.Axis':
            assert len(segments)==1 and abs(first[-1]-2*math.pi)<1e-12
            direction='Counterclockwise'
        if direction is None:return
        x1,y1=first[2:4];x2,y2=last[4:6]
        if direction=='Right':assert x2>x1 and abs(y2-y1)<1e-9,(case,path,first,last)
        if direction=='Left':assert x2<x1 and abs(y2-y1)<1e-9,(case,path,first,last)
        if direction=='Up':assert y2>y1 and abs(x2-x1)<1e-9,(case,path,first,last)
        if direction=='Down':assert y2<y1,(case,path,first,last)
        checked.append({'Case':case,'Path':path,'Direction':direction,'Start':[x1,y1],'End':[x2,y2]})
    for line in lines:
        if line.startswith('CASE|'):
            flush();case=line;path=None;segments=[]
        elif line.startswith('PATH|'):
            flush();path=line.split('|')[1];segments=[]
        elif line.startswith('SEGMENT|'):
            segments.append([struct.unpack('>d',bytes.fromhex(x))[0] for x in line.split('|')[5:]])
    flush()
    assert len({x['Case'] for x in checked})==60
    for name in ('B.Top','B.Bottom','H.Left','H.Right','Circle.Axis','Opening.B.Top','Opening.B.Bottom','Opening.H.Left','Opening.H.Right'):
        assert any(x['Path']==name for x in checked),name
    write('Directions.json',{'Passed':True,'RecipeCases':60,'CheckedPaths':len(checked),'Paths':checked})
    return len(checked)


def native_components(body):
    return {m[1]:m[2] for m in re.finditer(r'(?ms)^COMPONENT: ([^\n]+)\nTYPE: [^\n]+\nLINES: [^\n]+\n=+\n(.*?)(?=\n=+\nCOMPONENT: |\Z)',body)}


def native_equivalence():
    baseline=subprocess.check_output(['git','show',receipt('Before.json')['Head']+':workbook/output/VBA_All_Code.txt'],cwd=ROOT).decode('utf-8-sig').replace('\r\n','\n')
    a,b=native_components(baseline),native_components(text('CompiledNativeVBA.txt'))
    assert len(a)==len(b)==158 and a.keys()==b.keys()
    def normalized(body):
        return ''.join(p if p.startswith('"') else re.sub(r'\s+','',p.lower()) for p in re.findall(r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27]+',body) if not p.startswith("'"))
    for name in a:
        prepared,_=format_cases(a[name],name)
        assert normalized(prepared)==normalized(b[name]),name
    return {'Passed':True,'Components':len(a),'ExecutableVBAEquivalent':True,'AllowedCaseLayoutChanges':4}


def package():
    scope=source_scope();paths=directions()
    assert receipt('CodeUpdate.json')['Compiled'] and receipt('CompiledSourceEquality.json')['Passed']
    native=native_equivalence()
    with ZipFile(D/'Baseline.xlsm') as original,ZipFile(D/'Compiled.xlsm') as compiled,ZipFile(D/'Accepted.xlsm','w',ZIP_DEFLATED) as result:
        vba=compiled.read('xl/vbaProject.bin')
        for info in original.infolist():result.writestr(info,vba if info.filename=='xl/vbaProject.bin' else original.read(info.filename))
    with ZipFile(D/'Baseline.xlsm') as a,ZipFile(D/'Accepted.xlsm') as b:
        assert a.namelist()==b.namelist()
        assert [name for name in a.namelist() if a.read(name)!=b.read(name)]==['xl/vbaProject.bin']
    result={'Passed':True,'SourceScope':scope,'NativeEquivalence':native,'CheckedPaths':paths,'AllNonVBAPartsByteIdentical':True,
            'AcceptedSHA256':sha(D/'Accepted.xlsm'),'CompiledSHA256':sha(D/'Compiled.xlsm'),'CompiledVbaSHA256':hashlib.sha256(vba).hexdigest().upper(),
            'CalculationTestsRequired':False,'VisualRetestRequired':False}
    write('Package.json',result)
    return result


def accept(published):
    scope=source_scope();paths=directions();native=native_equivalence();package_data=receipt('Package.json')
    assert package_data['Passed'] and sha(D/'Accepted.xlsm')==package_data['AcceptedSHA256']
    assert sha(ROOT/'workbook/output/RC_Section_NDM_execution_report.txt')==receipt('Before.json')['UserReportSHA256']
    if published:
        assert sha(ROOT/'workbook/output/RC_Section_NDM.xlsm')==package_data['AcceptedSHA256']
        assert receipt('PublishedSourceEquality.json')['Passed']
        assert (ROOT/'workbook/output/VBA_All_Code.txt').read_bytes()==data('PublishedNativeVBA.txt')
        a,b=native_components(text('CompiledNativeVBA.txt')),native_components(text('PublishedNativeVBA.txt'))
        assert a==b,'Published code differs from compiled reference.'
        assert not re.search(r'NDM\.RECIPE\.GUIDE|[0-9a-fA-F]{64}',text('PublishedNativeVBA.txt'))
    result={'Passed':True,'PublishedVerified':published,'SourceScope':scope,'NativeEquivalence':native,'RecipeCases':60,'CheckedPaths':paths,
            'AllNonVBAPartsByteIdentical':True,'UserReportUnchanged':True}
    write('Acceptance.json',result)
    return result


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--package',action='store_true');parser.add_argument('--published',action='store_true')
    args=parser.parse_args()
    result=package() if args.package else accept(args.published)
    print(json.dumps({k:v for k,v in result.items() if k not in ('SourceScope',)},ensure_ascii=False))
