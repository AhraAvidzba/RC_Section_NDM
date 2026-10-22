"""Проверяет шапки VBA и фиксирует все исключения из размещения комментариев внутри строки."""
from pathlib import Path
import argparse, hashlib, json, re, subprocess

ROOT=Path(__file__).resolve().parents[2]
METHOD=re.compile(r'^\s*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)',re.I)
RUSSIAN=re.compile('[А-Яа-яЁё]')

def split_comment(line):
    """Не принимает апостроф внутри строки VBA за начало комментария."""
    quoted=False; i=0
    while i<len(line):
        if line[i]=='"':
            if quoted and i+1<len(line) and line[i+1]=='"': i+=2; continue
            quoted=not quoted
        elif line[i]=="'" and not quoted: return line[:i].rstrip(),line[i+1:].strip()
        i+=1
    return line.rstrip(),''

def executable(text):
    """Сохраняет строки, порядок операций и продолжения строк; убирает только комментарии."""
    out=[]
    for line in text.splitlines():
        code,_=split_comment(line)
        if code.strip() and not re.match(r'^\s*Rem(?:\s|$)',code,re.I): out.append(code.rstrip())
    return '\n'.join(out)

def preceding(lines,index):
    block=[]
    for line in reversed(lines[:index]):
        if not line.strip():
            if block: break
            continue
        if line.lstrip().startswith("'"): block.append(line)
        else: break
    return list(reversed(block))

def inspect(path):
    """Собирает структурные доказательства; смысловую проверку подтверждают отдельные отчёты ревизии."""
    text=path.read_text(encoding='utf-8-sig');lines=text.splitlines();methods=[];standalone=[];active=None
    for i,line in enumerate(lines):
        match=METHOD.match(line)
        if match:
            active={'Name':match[2],'Kind':match[1],'Line':i+1,'Header':preceding(lines,i)}
            active['HeaderRussian']=any(RUSSIAN.search(l) for l in active['Header'])
            active['TestMarker']=any('ДЛЯ ТЕСТОВ' in l for l in active['Header'])
            methods.append(active)
            if re.search(r':\s*End\s+(Sub|Function|Property)\b',split_comment(line)[0],re.I): active=None
        elif re.match(r'^\s*End\s+(Sub|Function|Property)\b',line,re.I): active=None
        elif active and line.lstrip().startswith("'"):
            j=i+1
            while j<len(lines) and (not lines[j].strip() or lines[j].lstrip().startswith("'")): j+=1
            target=lines[j].strip() if j<len(lines) else ''
            standalone.append({'Method':active['Name'],'Line':i+1,'Text':line.strip(),'NextCode':target})
    option=next((i for i,l in enumerate(lines) if l.strip()=='Option Explicit'),-1)
    start=option+1
    if option<0:
        start=0
        while start<len(lines) and (lines[start].startswith('Attribute ') or not lines[start].strip()): start+=1
    while start<len(lines) and (not lines[start].strip() or lines[start].startswith('Option ')): start+=1
    module_header=[]
    while start<len(lines) and lines[start].lstrip().startswith("'"):
        module_header.append(lines[start]);start+=1
    for method in methods:
        if method['Header']==module_header: method['HeaderRussian']=False
    is_test='tests' in path.parts or path.stem.startswith(('modTest','CTest','CFake')) or path.stem=='modTestFinalGeometryChecks'
    return {'Path':path.relative_to(ROOT).as_posix(),'Lines':len(lines),'ModuleHeader':module_header,
            'ModuleHeaderRussian':any(RUSSIAN.search(l) for l in module_header),'TestModule':is_test,
            'Methods':methods,'StandaloneInsideMethods':standalone,
            'ExecutableSHA256':hashlib.sha256(executable(text).encode()).hexdigest()}

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--report',type=Path,required=True)
    parser.add_argument('--base-ref');parser.add_argument('--check',action='store_true');args=parser.parse_args()
    files=sorted(p for d in ('src','tests','docs/geometry_refactor/examples') for p in (ROOT/d).rglob('*') if p.suffix in ('.cls','.bas'))
    files+=[ROOT/'tools/build_workbook/WorkbookHelpEvents.vba']
    modules=[inspect(p) for p in files]
    result={'Files':modules,'ModuleCount':len(modules),'MethodCount':sum(len(m['Methods']) for m in modules),
            'MissingModuleHeaders':[m['Path'] for m in modules if not m['ModuleHeaderRussian']],
            'MissingMethodHeaders':[{'Path':m['Path'],'Method':f['Name'],'Line':f['Line']} for m in modules for f in m['Methods'] if not f['HeaderRussian']],
            'MissingTestMarkers':[{'Path':m['Path'],'Method':f['Name'],'Line':f['Line']} for m in modules for f in m['Methods'] if m['TestModule'] and not f['TestMarker']]}
    exceptions_path=ROOT/'tools/build_workbook/VbaTestNameExceptions.json'
    exceptions=json.loads(exceptions_path.read_text(encoding='utf-8')) if exceptions_path.exists() else []
    allowed={(e['Module'],e['Method']) for e in exceptions}
    result['MissingTestNames']=[{'Path':m['Path'],'Method':f['Name'],'Line':f['Line']}
        for m in modules for f in m['Methods'] if (m['TestModule'] or f['TestMarker'])
        and 'Test' not in f['Name'] and (Path(m['Path']).stem,f['Name']) not in allowed]
    result['MissingTestModuleNames']=[m['Path'] for m in modules if m['TestModule'] and 'Test' not in Path(m['Path']).stem]
    result['MissingTestModuleMarkers']=[m['Path'] for m in modules if m['TestModule']
        and not any('ДЛЯ ТЕСТОВ' in h for h in m['ModuleHeader'])]
    result['MissingTestDividers']=[{'Path':m['Path'],'Method':f['Name'],'Line':f['Line']}
        for m in modules for f in m['Methods'] if (m['TestModule'] or f['TestMarker'])
        and not any('-------------------- ДЛЯ ТЕСТОВ --------------------' in h for h in f['Header'])]
    result['ContractNameExceptions']=len(allowed)
    result['ContractAccessorExceptions']=len(exceptions)
    if args.base_ref:
        changed=[]
        for path,module in zip(files,modules):
            old=subprocess.check_output(['git','show',args.base_ref+':'+module['Path']],cwd=ROOT).decode('utf-8-sig')
            if executable(old)!=executable(path.read_text(encoding='utf-8-sig')): changed.append(module['Path'])
        result['ChangedExecutableFiles']=changed
    args.report.parent.mkdir(parents=True,exist_ok=True)
    args.report.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({k:v for k,v in result.items() if k!='Files'},ensure_ascii=False))
    if args.check and any(v for k,v in result.items() if k.startswith('Missing')): raise SystemExit(1)

if __name__=='__main__':main()
