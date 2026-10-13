"""Приёмка всех рецептов: смысл аргументов, неизменные команды и сохранность книги."""
import hashlib
import importlib.util
import json
import re
from pathlib import Path

ROOT = Path.cwd()
directory = ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/RecipeArguments'
spec = importlib.util.spec_from_file_location('reader', ROOT/'tools/build_workbook/Verify-SectionRecipeGuide.py')
reader = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reader)
errors = []

def require(condition, reason):
    if not condition:
        errors.append(reason)

def split_comment(line):
    quoted, i = False, 0
    while i < len(line):
        if line[i] == '"':
            if quoted and i+1 < len(line) and line[i+1] == '"':
                i += 2
                continue
            quoted = not quoted
        elif line[i] == "'" and not quoted:
            return line[:i], line[i+1:].strip()
        i += 1
    return line, ''

def commands(text):
    result, pending = [], ''
    for line in text.splitlines():
        code, _ = split_comment(line)
        code = pending+code.strip()
        if code.endswith(' _'):
            pending = code[:-2]+' '
            continue
        pending = ''
        if code:
            parts = re.split(r'("(?:[^"\n]|"")*")', code)
            result.append(''.join(p if i % 2 else re.sub(r'\s+', '', p).lower() for i,p in enumerate(parts)))
    require(not pending, 'Unterminated VBA continuation')
    return result

def native(path):
    pattern = r'(?ms)^={100}\nCOMPONENT: ([^\n]+)\nTYPE: (\d+)\nLINES: (\d+)\n={100}\n(.*?)(?=^={100}\nCOMPONENT: |\Z)'
    return {m[1]:m[4].strip() for m in re.finditer(pattern, path.read_text(encoding='utf-8-sig'))}

manifest = json.loads((directory/'SourceComments.json').read_text(encoding='utf-8'))
recipes = sorted([p for base in ('src','docs/geometry_refactor/examples') for p in (ROOT/base).rglob('*.cls')
                  if re.search(r'(?m)^Implements ISectionRecipe\s*$',p.read_text(encoding='utf-8-sig'))])
require({str(p.relative_to(ROOT)).replace('\\','/') for p in recipes} == {p['source'] for p in manifest['recipes']}, 'Recipe inventory incomplete')
require(len(recipes) == 5, 'Expected four built-in recipes and one teaching recipe')
old_native, new_native = native(directory/'LiveNative.txt'), native(directory/'NativeVBA.txt')
example_native = native(directory/'ExampleFixture/NativeVBA.txt')
total, maximum = 0, 0
for recipe in manifest['recipes']:
    path = ROOT/recipe['source']
    before = (directory/'SourceBefore'/recipe['source']).read_text(encoding='utf-8-sig')
    after = path.read_text(encoding='utf-8-sig')
    require(commands(before) == commands(after), 'Executable source changed: '+path.stem)
    # Long explanations above a statement are explicitly allowed by the user.
    require(hashlib.sha256(path.read_bytes()).hexdigest() == recipe['sourceSHA256'], 'Source changed after comment inventory: '+path.stem)
    entries = {entry['line']:entry for entry in recipe['entries']}
    for number,line in enumerate(after.splitlines(),1):
        code, comment = split_comment(line)
        if number in entries:
            entry=entries[number]
            if entry['placement']=='inline':
                require(bool(code.strip()) and comment==entry['comment'], 'Missing inline explanation: '+path.stem+':'+str(number))
            else:
                notes=[]; previous=number-2
                while previous>=0 and after.splitlines()[previous].lstrip().startswith("'"):
                    notes.insert(0,after.splitlines()[previous].lstrip()[1:].strip()); previous-=1
                require(' '.join(notes).endswith(entry['comment']), 'Missing explanation above: '+path.stem+':'+str(number))
            require(entry['arguments'] or not re.search(r'facade\.(?!Finish)|options\.\w+\s*=',code), 'Arguments not explained: '+code)
            maximum=max(maximum,len(entry['comment'])); total+=1
        # Independent check of all builder calls and reinforcement option writes.
        creation = re.search(r'\bfacade\.|\boptions\.\w+\s*=|\bNew C(?:SectionDefinitionBuilder|RebarRuleOptions)\b', code)
        if creation and not code.lstrip().startswith(('Private ','Public ')):
            require(number in entries, 'Creation statement not covered: '+path.stem+':'+str(number))
    body = after[after.index('Option Explicit'):]
    exported = example_native if path.stem == 'CSectionRecipeRoundedTrapezoid' else new_native
    require(commands(body) == commands(exported[path.stem]), 'Recipe source/native commands differ: '+path.stem)
    native_text=' '.join(line.strip().lstrip("'").strip() for line in exported[path.stem].splitlines()); require(all(entry['comment'] in native_text for entry in recipe['entries']), 'Native comments incomplete: '+path.stem)
require(set(old_native) == set(new_native) and len(new_native) == 150, 'Production component inventory changed')
require(all(commands(old_native[name]) == commands(new_native[name]) for name in old_native), 'Published candidate executable code changed')
require(set(example_native)-set(new_native) == {'CSectionRecipeRoundedTrapezoid'}, 'Unexpected example fixture component')
guide = new_native['modSectionRecipeApiGuide'].splitlines()
start = guide.index("'   МИНИМАЛЬНЫЙ ПРИМЕР")
example = guide[start+1:next(i for i in range(start+1,len(guide)) if 'В CSectionTypeRegistry.Class_Initialize' in guide[i])]
example_lines = [line for line in example if line.startswith("'       ") and not line.startswith("'       '")]
require(len(example_lines) == 13, 'Expected thirteen commands in guide example')
require(all(split_comment(line[8:])[1] or example[example.index(line)-1].startswith("'       '") for line in example_lines), 'Guide command missing explanation')
# Above-line notes are permitted for long argument lists.
old, new = reader.read(directory/'LiveBefore.xlsm'), reader.read(directory/'Candidate.xlsm')
layout = json.loads((ROOT/'docs/geometry_refactor/SECTION_RECIPE_GUIDE_LAYOUT.json').read_text(encoding='utf-8'))
first = int(re.search(r'\$A\$(\d+)',old['names']['rngSectionRecipeGuide'])[1])
changes = {'A'+str(first+i):row['text'] for i,row in enumerate(layout['rows'])
           if old['sheets']['Справка']['cells'].get('A'+str(first+i),{}).get('value','') != row['text'] and row['text']}
require(len(changes) == 8, 'Unexpected Help edit scope')
require(old['names'] == new['names'], 'Named ranges changed')
require(old['drawings'] == new['drawings'], 'Drawings changed')
cells = 0
for name,data in old['sheets'].items():
    actual = new['sheets'][name]
    for address in data['cells'].keys() | actual['cells'].keys():
        a,b = data['cells'].get(address),actual['cells'].get(address)
        cells += 1
        if name == 'Справка' and address in changes:
            require(b['value'] == changes[address], 'Help text differs: '+address)
            require(all(a[k] == b[k] for k in ('formula','style','attrs')), 'Help formatting changed: '+address)
        else:
            require(a == b, 'Cell changed: '+name+'!'+address)
    require(data['rows'] == actual['rows'], 'Row metadata changed: '+name)
    for element in data['document']:
        if element.tag in (reader.S+'sheetData',reader.S+'sheetViews'):
            continue
        following = actual['document'].find(element.tag)
        if element.tag == reader.S+'mergeCells':
            require({e.get('ref') for e in element} == {e.get('ref') for e in following}, 'Merges changed: '+name)
        else:
            require(reader.canonical(element) == reader.canonical(following), 'Sheet structure changed: '+name+':'+element.tag)
compiled = json.loads((directory/'CodeUpdate.json').read_text(encoding='utf-8-sig'))['Compiled']
example_compiled = json.loads((directory/'ExampleFixture/CodeUpdate.json').read_text(encoding='utf-8-sig'))['Compiled']
require(compiled and example_compiled, 'All recipes must compile')
sha = lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
result = {'Passed':not errors,'RecipeClasses':len(recipes),'CommentedCommands':total,'ArgumentsExplained':manifest['argumentsExplained'],'Inline':manifest['inline'],'Above':manifest['above'],
          'GuideExampleInlineComments':len(example_lines),'AboveLineExceptionsAllowedByUser':True,'ExecutableCodeUnchanged':True,
          'ProductionCompiled':compiled,'ExampleCompiled':example_compiled,'NativeComponents':len(new_native),
          'CellsChecked':cells,'NamedRanges':len(old['names']),'ChangedHelpCells':sorted(changes),
          'CandidateSHA256':sha(directory/'Candidate.xlsm'),'NativeSHA256':sha(directory/'NativeVBA.txt'),'Errors':errors}
(directory/'Preservation.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(result,ensure_ascii=False))
raise SystemExit(bool(errors))
