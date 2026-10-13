"""Статическая инвентаризация VBA и чтение кода с сохранением номеров строк."""
import argparse
import json
import re
from pathlib import Path

root = Path.cwd()

def code_only(line):
    quoted, i = False, 0
    while i < len(line):
        if line[i] == '"':
            if quoted and i+1 < len(line) and line[i+1] == '"':
                i += 2
                continue
            quoted = not quoted
        elif line[i] == "'" and not quoted:
            return line[:i].rstrip()
        i += 1
    return line.rstrip()

parser = argparse.ArgumentParser()
parser.add_argument('files', nargs='*')
parser.add_argument('--first', type=int, default=1)
parser.add_argument('--last', type=int, default=100000)
parser.add_argument('--inventory', type=Path)
args = parser.parse_args()
if args.inventory:
    sources = sorted((root/'src').rglob('*.cls'))
    by_name = {p.stem: p for p in sources}
    result = []
    for path in sources:
        text = path.read_text(encoding='utf-8-sig')
        code = '\n'.join(code_only(line) for line in text.splitlines())
        deps = sorted(set(re.findall(r'\b(?:As|New|Implements)\s+(C\w+|I\w+)\b',code)) & by_name.keys())
        methods = [{'line':i,'declaration':line.strip()} for i,line in enumerate(text.splitlines(),1)
                   if re.match(r'^(Public|Friend|Private) (Sub|Function|Property)',line)]
        result.append({'class':path.stem,'path':path.as_posix().removeprefix(root.as_posix()+'/'),
                       'lines':len(text.splitlines()),'dependencies':deps,'methods':methods})
    args.inventory.parent.mkdir(parents=True,exist_ok=True)
    args.inventory.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    for item in result:
        if '/Geometry/' in item['path'] or '/Section/' in item['path']:
            print(item['class'],item['lines'],'->',', '.join(item['dependencies']))
else:
    for name in args.files:
        path = root/name
        print('\n'+name)
        for i,line in enumerate(path.read_text(encoding='utf-8-sig').splitlines(),1):
            code = code_only(line)
            if args.first<=i<=args.last and code.strip() and not code.startswith(('VERSION ','BEGIN','  MultiUse','END','Attribute ')):
                print(f'{i:4}: {code}')
