"""Приёмка разделения чтения и построения в четырёх штатных рецептах."""
import argparse
import hashlib
import importlib.util
import json
import re
import subprocess
from collections import Counter
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parents[2]
D = ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/RecipeResponsibilities_2026-10-16'
RECIPES = ['CSectionRecipe'+x for x in ('Circle','RectSet','PointedRect','HollowRectangle')]
MATRICES = ['modTestSectionRecipes.Dump'+x+'Recipe'+suffix for x,suffix in
            [('Circle','Matrix'),('RectSet','Matrix'),('Rounded','Matrix'),('Hollow','Matrix'),('Hollow','EdgeMatrix')]]
CONTRACTS = ['modTestAudit02Snapshots.DumpRecipeRuleContracts','modTestAudit02Api.DumpAudit02CloseoutInputContract']
STATES = ['modTestSectionRecipes.Dump'+x+'RecipeCalculations' for x in ('Circle','RectSet','Rounded','Hollow')]
PATHS = 'modTestAudit02Snapshots.DumpRecipePathContracts'
spec = importlib.util.spec_from_file_location('guide', Path(__file__).with_name('Verify-SectionRecipeGuide.py'))
guide = importlib.util.module_from_spec(spec)
spec.loader.exec_module(guide)


def data(name):
    if (D/name).exists(): return (D/name).read_bytes()
    if not (D/'Evidence.zip').exists(): raise FileNotFoundError(D/name)
    with ZipFile(D/'Evidence.zip') as z: return z.read(name)


def text(name): return data(name).decode('utf-8-sig').replace('\r\n','\n')
def receipt(name): return json.loads(text(name))
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def write(name, value):
    (D/name).write_text(json.dumps(value,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')


def scope():
    expected = {'src/Geometry/'+x+'.cls' for x in RECIPES} | {'src/Common/modSectionRecipeApiGuide.bas','tests/modTestSectionRecipes.bas'}
    with ZipFile(D/'BeforeSources.zip') as z:
        before = set(z.namelist())
        after = {p.relative_to(ROOT).as_posix() for folder in ('src','tests') for p in (ROOT/folder).rglob('*') if p.suffix in ('.cls','.bas')}
        assert before == after, 'A source component was added or removed.'
        changed = {name for name in before if z.read(name).decode('utf-8-sig').replace('\r\n','\n') != (ROOT/name).read_text(encoding='utf-8-sig')}
        assert changed == expected, changed
        original = z.read('tests/modTestSectionRecipes.bas').decode('utf-8-sig').replace('\r\n','\n').rstrip()
        current = (ROOT/'tests/modTestSectionRecipes.bas').read_text(encoding='utf-8-sig')
        assert current.startswith(original+'\n'), 'Existing tests changed.'
    for name in RECIPES:
        source = (ROOT/'src/Geometry'/(name+'.cls')).read_text(encoding='utf-8-sig')
        assert 'RowByAnyKey' not in source
        assert all(x in source for x in ('ReadRebarInputs','AddGeometry','AddReinforcement','AddAnnotations'))
        for method,body in re.findall(r'(?ms)^Private (?:Sub|Function) (\w+)\((.*?)^End (?:Sub|Function)',source):
            if method.startswith('Add') or method in ('AxisPath','DescribePolygon'):
                assert not re.search(r'CSectionInputTable|CUnitSystem|\btable\.|\bcontext\.|\bApplication\b|\bWorksheet\b',body), (name,method)
            if method.startswith('Read'):
                assert 'facade' not in body and 'New CGeometryPath' not in body, (name,method)
    assert receipt('CommentAudit.json')['status'] == 'PASS'
    return sorted(changed)


def native_scope():
    before=subprocess.check_output(['git','show',receipt('Before.json')['Head']+':workbook/output/VBA_All_Code.txt'],cwd=ROOT).decode('utf-8-sig').replace('\r\n','\n')
    after=text('AcceptedNativeVBA.txt')
    def components(body):
        result={}
        for match in re.finditer(r'(?ms)^COMPONENT: ([^\n]+)\nTYPE: [^\n]+\nLINES: [^\n]+\n=+\n(.*?)(?=\n=+\nCOMPONENT: |\Z)',body):
            result[match[1]]=match[2]
        return result
    def executable(body):
        return ''.join(p if p.startswith('"') else re.sub(r'\s+','',p.lower()) for p in
                       re.findall(r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27]+',body) if not p.startswith("'"))
    a,b=components(before),components(after)
    assert len(a)==len(b)==158 and a.keys()==b.keys(),(len(a),len(b))
    changed=sorted(name for name in a if executable(a[name])!=executable(b[name]))
    assert set(changed)==set(RECIPES+['modTestSectionRecipes']),changed
    return {'Components':len(a),'ChangedExecutableComponents':changed,'OtherExecutableComponentsUnchanged':len(a)-len(changed)}


def run(stage, expected=None):
    record = receipt(stage+'/Run.json')
    assert record['Passed'] and record['ProductionUnchanged'], stage
    names = [r['Macro'] for r in record['Runs']]
    if expected is not None: assert names == expected, (stage,names)
    count = 0
    for item in record['Runs']:
        body = text(stage+'/'+item['Macro']+'.txt')
        assert item['Passed'] and body.strip()
        assert not re.search(r'(?m)^FAIL[: ]|failed=[1-9]|RUNTIME ERROR:|^ERROR:',body), stage
        totals = re.findall(r'(?m)^TOTAL[^\n]*passed=(\d+); failed=0\b',body)
        if not item['Snapshot'] or 'RunRecipeCaptionTests' in item['Macro']:
            assert totals, stage
            count += int(totals[-1])
    return {'Stage':stage,'Assertions':count,'Seconds':sum(x['Seconds'] for x in record['Runs'])}


def tests():
    comparisons = []
    stages = [run('BaselineMatrices',MATRICES+CONTRACTS),run('CandidateMatrices',MATRICES+CONTRACTS),run('Directed'),
              run('BaselinePaths',[PATHS]),run('CompiledContracts',['modTestSectionRecipes.RunRecipeCaptionTests',PATHS])]
    pairs = [('BaselineMatrices','CandidateMatrices',MATRICES+CONTRACTS),('BaselinePaths','CompiledContracts',[PATHS])]
    for suffix in ('','63'):
        stages += [run('BaselineStates'+suffix,STATES),run('CandidateStates'+suffix,STATES)]
        pairs.append(('BaselineStates'+suffix,'CandidateStates'+suffix,STATES))
    for before,after,macros in pairs:
        for macro in macros:
            name = macro+'.txt'
            assert data(before+'/'+name) == data(after+'/'+name), (before,after,macro)
            comparisons.append({'Before':before+'/'+name,'After':after+'/'+name,'SHA256':hashlib.sha256(data(after+'/'+name)).hexdigest()})
    rows = [line.split('\t') for suffix in ('','63') for macro in STATES for line in text('CandidateStates'+suffix+'/'+macro+'.txt').splitlines()]
    nds = [row for row in rows if row[1]=='NDS']
    assert len(nds)==240 and all(row[3]=='b:1' for row in nds)
    assert len([r for r in rows if r[1]=='BATCH'])==40
    assert text('CandidateMatrices/'+CONTRACTS[0]+'.txt').count('CASE|')==60
    assert receipt('CompiledSourceEquality.json')['Passed']
    return {'DirectedAssertions':sum(s['Assertions'] for s in stages),'ExactComparisons':comparisons,
            'RuleCases':60,'NDSStates':len(nds),'BatchCombinations':40,'StateRows':dict(Counter(r[1] for r in rows)), 'Stages':stages}


def package():
    scope()
    result = tests()
    assert receipt('CodeUpdate.json')['Compiled']
    before = guide.read(D/'Baseline.xlsm')
    old = before['sheets']['Справка']['cells']['A611']['value']
    new = (ROOT/'docs/geometry_refactor/ADDING_SECTION_TYPE.md').read_text(encoding='utf-8-sig').strip().splitlines()[-1]
    assert old.startswith('Как примеры нового сечения') and new.startswith('Как примеры нового сечения') and old != new
    with ZipFile(D/'Compiled.xlsm') as compiled, ZipFile(D/'Baseline.xlsm') as original, ZipFile(D/'Accepted.xlsm','w',ZIP_DEFLATED) as output:
        vba = compiled.read('xl/vbaProject.bin')
        for info in original.infolist():
            body = original.read(info.filename)
            if info.filename=='xl/vbaProject.bin': body=vba
            if info.filename=='xl/sharedStrings.xml':
                old_bytes,new_bytes=escape(old).encode('utf-8'),escape(new).encode('utf-8')
                assert body.count(old_bytes)==1
                body=body.replace(old_bytes,new_bytes)
            output.writestr(info,body)
    after = guide.read(D/'Accepted.xlsm')
    assert before['names']==after['names']
    assert before['drawings']==after['drawings']
    cells=0
    for sheet,a in before['sheets'].items():
        b=after['sheets'][sheet]
        assert a['rows']==b['rows'] and a['cells'].keys()==b['cells'].keys()
        for address,value in a['cells'].items():
            expected=dict(value)
            if sheet=='Справка' and address=='A611': expected['value']=new
            assert expected==b['cells'][address],(sheet,address)
            cells+=1
        assert guide.canonical(a['document'])==guide.canonical(b['document']),sheet
    with ZipFile(D/'Baseline.xlsm') as a,ZipFile(D/'Accepted.xlsm') as b:
        assert a.namelist()==b.namelist()
        changed=[name for name in a.namelist() if a.read(name)!=b.read(name)]
        assert sorted(changed)==['xl/sharedStrings.xml','xl/vbaProject.bin'],changed
    result.update(Passed=True,ChangedComponents=scope(),ChangedZIPParts=changed,PreservedCells=cells,PreservedNames=len(before['names']),
                  AcceptedSHA256=sha(D/'Accepted.xlsm'),CompiledSHA256=sha(D/'Compiled.xlsm'),
                  CompiledVbaSHA256=hashlib.sha256(vba).hexdigest().upper(),HelpChange={'Cell':'Справка!A611','Before':old,'After':new})
    write('Package.json',result)
    return result


def accept(published=False):
    result=tests()
    package_data=receipt('Package.json')
    assert package_data['Passed'] and sha(D/'Accepted.xlsm')==package_data['AcceptedSHA256']
    assert receipt('AcceptedSourceEquality.json')['Passed']
    run('AcceptedSmoke',['modTestSectionRecipes.RunWorkbookRecipeInputTests','modTestSectionRecipes.RunRecipeCaptionTests'])
    with ZipFile(D/'Accepted.xlsm') as z:
        assert hashlib.sha256(z.read('xl/vbaProject.bin')).hexdigest().upper()==package_data['CompiledVbaSHA256']
        marker = re.compile(r'NDM\.RECIPE\.GUIDE|[0-9a-fA-F]{64}')
        # Hashes belong in external receipts, not in worksheet text or VBA literals.
        assert not marker.search(text('AcceptedNativeVBA.txt'))
        for name in z.namelist():
            if name.endswith('.xml'): assert not marker.search(z.read(name).decode('utf-8-sig')),name
    assert receipt('HelpReview.json')['Passed']
    assert sha(ROOT/'workbook/output/RC_Section_NDM_execution_report.txt')==receipt('Before.json')['UserReportSHA256']
    if published:
        assert sha(ROOT/'workbook/output/RC_Section_NDM.xlsm')==package_data['AcceptedSHA256']
        assert receipt('PublishedSourceEquality.json')['Passed']
        assert (ROOT/'workbook/output/VBA_All_Code.txt').read_bytes()==data('PublishedNativeVBA.txt')
    result.update(Passed=True,PublishedVerified=published,ChangedComponents=scope(),NativeScope=native_scope(),PreservedCells=package_data['PreservedCells'],PreservedNames=package_data['PreservedNames'])
    write('Acceptance.json',result)
    return result


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--package',action='store_true')
    parser.add_argument('--published',action='store_true')
    args=parser.parse_args()
    result=package() if args.package else accept(args.published)
    print(json.dumps({k:v for k,v in result.items() if k not in ('ExactComparisons','HelpChange')},ensure_ascii=False))
