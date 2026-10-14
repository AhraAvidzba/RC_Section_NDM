"""Сверяет полные учебные классы/тесты с реально исполненным VBProject fixture."""
from pathlib import Path
import hashlib
import json
import re

ROOT=Path(__file__).resolve().parents[2]
S=ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14'

def commands(text):
    result=[]
    for line in text.splitlines():
        parts=re.split(r'("(?:[^"\n]|"")*"|\x27[^\n]*)',line)
        code=''.join(token if i%2 else re.sub(r'\s+','',token.lower())
                     for i,token in enumerate(parts) if not token.startswith("'"))
        if code:result.append(code)
    return result

def verify():
    export=(S/'ExampleNativeVBA.txt').read_text(encoding='utf-8-sig')
    pattern=r'(?ms)^={100}\nCOMPONENT: ([^\n]+)\nTYPE: (\d+)\nLINES: (\d+)\n={100}\n(.*?)(?=^={100}\nCOMPONENT: |\Z)'
    native={m[1]:m[4] for m in re.finditer(pattern,export)}
    rows=[]
    for name,extension in [('CSectionRecipeMySection','cls'),('CSectionRecipeRoundedTrapezoid','cls'),('modTestRoundedTrapezoidExample','bas')]:
        source=ROOT/'docs/geometry_refactor/examples'/(name+'.'+extension)
        text=source.read_text(encoding='utf-8-sig')
        if extension=='cls':text=text[text.index('Option Explicit'):]
        else:text='\n'.join(text.splitlines()[1:])
        rows.append({'Component':name,'ExecutableEqual':commands(text)==commands(native[name]),
                     'SourceSHA256':hashlib.sha256(source.read_bytes()).hexdigest()})
    registration='RegisterRecipe "RoundedTrapezoid", New CSectionRecipeRoundedTrapezoid, Array("rngRoundedTrapezoidGeometry")'
    registered=commands(native['CSectionTypeRegistry']).count(commands(registration)[0])==1
    run=json.loads((S/'ExampleTests/Run.json').read_text(encoding='utf-8-sig'))
    fixture_hash=hashlib.sha256((S/'ExampleFixture.xlsm').read_bytes()).hexdigest()
    passed=all(row['ExecutableEqual'] for row in rows) and registered and run['Passed'] and run['SourceSHA256'].lower()==fixture_hash
    result={'Passed':passed,'Rows':rows,'OneRegistration':registered,'NativeComponents':len(native),'RunPassed':run['Passed'],
            'FixtureSHA256':fixture_hash,'Checks':'94 build, 250 input, 5 minimal through ISectionRecipe. Code compared with literals unchanged, ignoring only comments/whitespace/VBE code case.'}
    (S/'ExampleEquality.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False))
    if not passed:raise SystemExit(1)

if __name__=='__main__':verify()
