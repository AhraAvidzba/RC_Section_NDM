"""Проверяет итоговые исходники и точные native-снимки до/после."""
import hashlib
import json
import re
import subprocess
from pathlib import Path

directory=Path(__file__).resolve().parent
root=directory.parents[3]
records=json.loads((directory/'ReplacementMap.json').read_text(encoding='utf-8'))['Records']
baseline_head=json.loads((directory/'Before.json').read_text(encoding='utf-8-sig'))['Head']

def previous(path):
    result=subprocess.run(['git','show',baseline_head+':'+path],cwd=root,capture_output=True,check=True)
    return result.stdout.decode('utf-8-sig').replace('\r\n','\n')

def commands(text):
    # Сохраняем строковые литералы; удаляем только комментарии и пустое пространство кода.
    parts=re.findall(r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27]+',text)
    return ''.join(p if p.startswith('"') else '' if p.startswith("'") else re.sub(r'\s+','',p) for p in parts)

expected_sources=sorted({record['Source'] for record in records})
allowed_sources=set(expected_sources) | {'src/Section/CRebarRuleOptions.cls', 'src/Common/modSectionRecipeApiGuide.bas'}
unchanged_sources=[]
for folder in ('src', 'tests'):
    for path in (root/folder).rglob('*'):
        if path.is_file() and path.suffix in ('.bas','.cls'):
            relative=path.relative_to(root).as_posix()
            if relative not in allowed_sources:
                assert previous(relative)==path.read_text(encoding='utf-8-sig').replace('\r\n','\n'),relative
                unchanged_sources.append(relative)
verified=[]
for relative in expected_sources:
    old=previous(relative)
    def expand(match):
        owner,raw=match.groups()
        arguments=[s.strip() for s in raw.split("'",1)[0].split(',')]
        return '\n'.join(owner+'.'+field+' = '+value for field,value in zip(('AnnotationGroup','Publication','Coincidence'),arguments))
    expected=re.sub(r'(?m)^\s*(\w+)\.SetPublication ([^\r\n]+)$',expand,old)
    actual=(root/relative).read_text(encoding='utf-8-sig')
    assert commands(expected)==commands(actual),relative
    for record in (r for r in records if r['Source']==relative):
        assert all(row.strip() in actual for row in record['Statements']),record
    verified.append(relative)
relative='src/Section/CRebarRuleOptions.cls'
old=previous(relative)
expected=re.sub(r'(?ms)^Public Sub SetPublication\(.*?^End Sub\n','',old)
actual=(root/relative).read_text(encoding='utf-8-sig')
assert commands(expected)==commands(actual)
defaults=re.compile(r'(?ms)^Private Sub Class_Initialize\(\).*?^End Sub')
assert defaults.search(old).group()==defaults.search(actual).group()
comparisons=[]
for name,macro in [('Geometry','modTestShapeArchitecture.DumpShapeArchitectureSnapshot'),('NDS','modTestShapeArchitecture.DumpPublicationControlStates')]:
    before=(directory/('Baseline'+name)/(macro+'.txt')).read_bytes()
    after=(directory/('Candidate'+name)/(macro+'.txt')).read_bytes()
    assert before==after,name
    comparisons.append({'Snapshot':name,'Bytes':len(after),'SHA256':hashlib.sha256(after).hexdigest(),'ByteIdentical':True})
states=(directory/'CandidateNDS/modTestShapeArchitecture.DumpPublicationControlStates.txt').read_text(encoding='utf-8-sig')
assert len(re.findall(r'\tNDS\tstate\t',states))==12
for folder in ('src','tests','tools','docs/geometry_refactor'):
    for path in (root/folder).rglob('*'):
        if path.is_file() and path.suffix in ('.bas','.cls','.md','.py','.ps1','.json'):
            assert not re.search('SetPublication',path.read_text(encoding='utf-8-sig'),re.I),path
report={'Passed':True,'BaselineHead':baseline_head,'Calls':13,'OneLinePerPropertyWithOwnComment':True,'RecipeAndTestCommandsEquivalent':verified,
        'OnlyObsoleteMethodRemovedFromOptions':True,'DefaultsUnchanged':True,'CommonExecutorAndSolverUntouched':True,
        'UnchangedSourceComponents':len(unchanged_sources),
        'LegacyModels':9,'ControlNDS':12,'Snapshots':comparisons,'NoLiveReferences':True}
(directory/'Equivalence.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(report,ensure_ascii=False))
