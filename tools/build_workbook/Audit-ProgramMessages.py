"""Lists executable message text in every VBA component for a complete review."""
import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[2]
METHOD = re.compile(r'^\s*(?:(?:Public|Private|Friend|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)', re.I)
LITERAL = re.compile(r'"(?:[^"\n]|"")*"')
CYRILLIC = re.compile('[А-Яа-яЁё]')

def executable(line):
    quoted = False
    i = 0
    while i < len(line):
        if line[i] == '"':
            if quoted and i + 1 < len(line) and line[i + 1] == '"':
                i += 2
                continue
            quoted = not quoted
        elif line[i] == "'" and not quoted:
            return line[:i].rstrip()
        i += 1
    return line.rstrip()

def statements(path):
    pending = ''
    first = 0
    method = '(declarations)'
    for number, line in enumerate(path.read_text(encoding='utf-8-sig').splitlines(), 1):
        code = executable(line)
        match = METHOD.match(code)
        if match:
            method = match[1]
        if not pending:
            first = number
        continued = code.endswith(' _')
        pending += (' ' if pending else '') + (code[:-2] if continued else code).strip()
        if continued:
            continue
        if pending:
            yield first, method, pending
        pending = ''

def audit(output):
    output.mkdir(parents=True, exist_ok=True)
    components = []
    areas = {}
    for path in sorted(ROOT.joinpath('src').rglob('*')):
        if path.suffix not in ('.cls', '.bas'):
            continue
        relative = path.relative_to(ROOT).as_posix()
        records = []
        for line, method, code in statements(path):
            literals = LITERAL.findall(code)
            if not any(CYRILLIC.search(text) for text in literals) and not re.search(r'\bErr\.Raise\b', code, re.I):
                continue
            records.append({'Line': line, 'Method': method, 'Code': code,
                            'RaisesError': bool(re.search(r'\bErr\.Raise\b', code, re.I))})
        components.append({'File': relative, 'SHA256': hashlib.sha256(path.read_bytes()).hexdigest().upper(),
                           'Statements': records, 'Count': len(records)})
        area = path.parent.name
        lines = areas.setdefault(area, [])
        lines.append(f'\n=== {relative}: {len(records)} statements ===')
        lines.extend(f'{r["Line"]} [{r["Method"]}] {r["Code"]}' for r in records)
    result = {'Components': len(components), 'Classes': sum(c['File'].endswith('.cls') for c in components),
              'Statements': sum(c['Count'] for c in components),
              'Raises': sum(r['RaisesError'] for c in components for r in c['Statements']),
              'ByArea': dict(Counter(c['File'].split('/')[1] for c in components)), 'Files': components}
    (output / 'Inventory.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    for area, lines in areas.items():
        (output / f'{area}.txt').write_text('\n'.join(lines) + '\n', encoding='utf-8')
    print(json.dumps({key: value for key, value in result.items() if key != 'Files'}, ensure_ascii=False))

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, required=True)
    audit(parser.parse_args().output)
