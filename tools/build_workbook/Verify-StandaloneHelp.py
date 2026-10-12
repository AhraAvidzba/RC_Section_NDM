"""Проверяет самостоятельность Справки и фактический модуль API в сохранённой книге."""
import argparse
import hashlib
import importlib.util
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('reader', Path(__file__).with_name('Verify-SectionRecipeGuide.py'))
reader = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reader)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--book', type=Path, required=True)
    parser.add_argument('--dump', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    args = parser.parse_args()
    book = reader.read(args.book)
    export = args.dump.read_text(encoding='utf-8-sig')
    pattern = r'(?ms)^={100}\r?\nCOMPONENT: ([^\r\n]+)\r?\nTYPE: (\d+)\r?\nLINES: (\d+)\r?\n={100}\r?\n(.*?)(?=^={100}\r?\nCOMPONENT: |\Z)'
    components = {item[1]: item[4] for item in re.finditer(pattern, export)}
    native = components['modSectionRecipeApiGuide'].replace('\r\n', '\n').strip()
    source = (ROOT/'src/Common/modSectionRecipeApiGuide.bas').read_text(encoding='utf-8-sig')
    source = '\n'.join(source.splitlines()[1:]).strip()
    errors = []
    if native != source: errors.append('Native API comments differ from source')
    if any(line.strip() and not line.lstrip().startswith("'") for line in native.splitlines()):
        errors.append('API module contains executable code')
    # Имена методов сохраняются; их описание не требует поставки исходников или отчётов.
    references = re.compile(r'(?:\b(?:docs|src|tests|tools|reference|norms)/|[\w-]+\.(?:md|txt|json|cls|bas|ps1|py|pdf)\b|\bфайл[а-яё]*\b|\bдокументац[а-яё]*\s+проекта\b)', re.I)
    checked = 0
    for sheet in ('Справка', 'Config'):
        for address, cell in book['sheets'][sheet]['cells'].items():
            text = cell['value']
            if not text: continue
            checked += 1
            if references.search(text): errors.append(sheet+'!'+address+': external implementation file reference: '+text)
    if references.search(native): errors.append('API comments refer to an external implementation file')
    if re.search(r'станци[а-яё]*', native, re.I): errors.append('Unexplained Russian station terminology remains')
    flattened = re.sub(r'\s+', ' ', '\n'.join(line.lstrip()[1:].strip() for line in native.splitlines()))
    catalog = json.loads((ROOT/'docs/geometry_refactor/SECTION_RECIPE_API_MANIFEST.json').read_text(encoding='utf-8-sig'))
    for member in catalog['members']:
        if re.sub(r'\s+', ' ', member['signature']) not in flattened:
            errors.append('API signature missing: '+member['id'])
    result = {'status': 'STANDALONE_HELP_PASS' if not errors else 'FAIL', 'cellsChecked': checked,
              'apiMembers': len(catalog['members']), 'commentOnly': True, 'nativeCommentsMatchSource': native == source,
              'workbookSHA256': hashlib.sha256(args.book.read_bytes()).hexdigest(),
              'dumpSHA256': hashlib.sha256(args.dump.read_bytes()).hexdigest(), 'errors': errors}
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps(result, ensure_ascii=False))
    if errors: raise SystemExit(1)


if __name__ == '__main__': main()
