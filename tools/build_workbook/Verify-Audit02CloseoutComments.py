"""Адресные guards комментариев: контракт helper, полная команда VBA и прежний исполняемый код."""
from pathlib import Path
import json
import re
import subprocess
from RecipeArgumentText import argument_comment, signatures, call_args
from SectionRecipeApiText import parameter_meaning

ROOT=Path(__file__).resolve().parents[2]
S=ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/Audit02Closeout_2026-10-14'

def commands(text):
    result=[]
    for line in text.splitlines():
        tokens=re.split(r'("(?:[^"\n]|"")*"|\x27[^\n]*)',line)
        code=''.join(token if i%2 else re.sub(r'\s+','',token.lower()) for i,token in enumerate(tokens) if not token.startswith("'"))
        if code:result.append(code)
    return result

def verify():
    checks={}
    baseline=json.loads((S/'Before.json').read_text(encoding='utf-8-sig'))['Head']
    files=[p for p in (ROOT/'src').rglob('*') if p.suffix in ('.cls','.bas')]
    for path in files:
        old=subprocess.check_output(['git','show',baseline+':'+path.relative_to(ROOT).as_posix()],cwd=ROOT).decode('utf-8-sig')
        checks[path.stem+'.executable.unchanged']=commands(old)==commands(path.read_text(encoding='utf-8-sig'))
    for name in ('baseDiameter','diameter'):
        meaning=parameter_meaning({'owner':'CRebarRuleOptions','name':'ConfigureDependentContact'}, {'name':name,'passing':'ByVal'})
        checks['contact.'+name+'.strict']= '>0' in meaning and 'ошибку' in meaning and 'до вызова' in meaning and 'выключает' not in meaning
    normal=parameter_meaning({'owner':'CSectionDefinitionBuilder','name':'RebarByCount'}, {'name':'diameter','passing':'ByVal'})
    checks['row.diameter.zero.disables']= 'ноль выключает ряд' in normal
    source=(ROOT/'src/Geometry/CSectionRecipePointedSection.cls').read_text(encoding='utf-8-sig')
    call=re.search(r'(?m)^    facade.Outer[^\n]+\n[^\n]+',source)[0]
    sig=signatures((ROOT/'src/Section/CSectionDefinitionBuilder.cls').read_text(encoding='utf-8-sig'))
    description,args=argument_comment(call.strip(),'Rounded',sig,{},'')
    checks['continuation.not.parameter']=all(a['value']!='_' for a in args) and '_ —' not in description
    checks['named.arguments.complete']= [a['parameter'] for a in args]==['loopID','pathID','comment','arcComment','namingPolicy']
    try:call_args('facade.Outer "OUTER", path.ID, "comment", _','Outer','facade.')
    except ValueError:checks['unfinished.command.rejected']=True
    else:checks['unfinished.command.rejected']=False
    for name in ('CSectionRecipePointedSection','CSectionRecipeHollowRectangle'):
        text=(ROOT/'src/Geometry'/(name+'.cls')).read_text(encoding='utf-8-sig')
        checks[name+'.no.stale.comment']=not re.search(r'_ —|sizes — размеры, вершины',text)
    guide=(ROOT/'src/Common/modSectionRecipeApiGuide.bas').read_text(encoding='utf-8-sig')
    checks['guide.comment.only']=all(not line.strip() or line.lstrip().startswith("'") for line in guide.splitlines()[1:])
    checks['guide.continuation.explained']='Это знак записи кода' in guide and 'не передаваемый параметр' in guide
    result={'Passed':all(checks.values()),'ProductionExecutableComponentsChecked':len(files),'Checks':checks}
    (S/'CommentContracts.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({'Passed':result['Passed'],'Checks':len(checks),'ProductionComponents':len(files),'Failed':[k for k,v in checks.items() if not v]},ensure_ascii=False))
    if not result['Passed']:raise SystemExit(1)

if __name__=='__main__':verify()
