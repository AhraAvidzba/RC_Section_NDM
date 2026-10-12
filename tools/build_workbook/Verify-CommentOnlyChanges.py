"""Доказывает неизменность VBA-токенов до и после дополнения комментариев."""
import argparse
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]

def executable_text(text):
    output = []
    for line in text.splitlines():
        quoted, index = False, 0
        while index < len(line):
            if line[index] == '"':
                if quoted and index+1 < len(line) and line[index+1] == '"':
                    index += 2
                    continue
                quoted = not quoted
            elif line[index] == "'" and not quoted: break
            index += 1
        code = line[:index].strip()
        if code and not code.lower().startswith('rem '): output.append(code)
    return '\n'.join(output)

def snapshot():
    return {p.relative_to(ROOT).as_posix(): hashlib.sha256(executable_text(p.read_text(encoding='utf-8-sig')).encode()).hexdigest()
            for folder in ('src', 'tests', 'docs/geometry_refactor/examples') for p in (ROOT/folder).rglob('*.cls')}

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--before', type=Path, required=True)
    parser.add_argument('--capture', action='store_true')
    parser.add_argument('--report', type=Path)
    args = parser.parse_args()
    current = snapshot()
    if args.capture:
        args.before.write_text(json.dumps(current, indent=2)+'\n', encoding='utf-8')
        return
    before = json.loads(args.before.read_text(encoding='utf-8'))
    differences = [p for p in set(before) | set(current) if before.get(p) != current.get(p)]
    result = {'status': 'PASS' if not differences else 'FAIL', 'classes': len(current), 'changedExecutableClasses': differences}
    if args.report: args.report.write_text(json.dumps(result, indent=2)+'\n', encoding='utf-8')
    print(json.dumps(result))
    if differences: raise SystemExit(1)

if __name__ == '__main__': main()
