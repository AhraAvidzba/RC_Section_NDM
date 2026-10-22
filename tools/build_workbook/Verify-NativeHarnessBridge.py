"""Prove that the release differs only in one private native-export test."""
import argparse
import hashlib
import json
import re
from pathlib import Path

PATTERN = r'(?ms)^={100}\r?\nCOMPONENT: ([^\r\n]+)\r?\nTYPE: (\d+)\r?\nLINES: (\d+)\r?\n={100}\r?\n(.*?)(?=^={100}\r?\nCOMPONENT: |\Z)'
BLOCK = r'(?ms)^Private Sub TestCheckNativeFullExport\(.*?^End Sub\b'


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def normalize(body):
    body = body.replace('\r\n', '\n').rstrip()
    start = body.find('Option Explicit')
    if start >= 0:
        body = body[start:]
    output = []
    for line in body.splitlines():
        quoted = False
        for index, char in enumerate(line):
            if char == '"':
                quoted = not quoted
                output.append(char)
            elif char == "'" and not quoted:
                output.append(line[index:])
                break
            else:
                output.append(char if quoted else char.lower())
        output.append('\n')
    return ''.join(output).rstrip()


def read(path):
    return {m[1]: (int(m[2]), m[4]) for m in re.finditer(PATTERN, path.read_text(encoding='utf-8-sig'))}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ('before-book', 'after-book', 'before-dump', 'after-dump', 'report'):
        parser.add_argument('--' + name, type=Path, required=True)
    args = parser.parse_args()
    before, after = read(args.before_dump), read(args.after_dump)
    errors, changed = [], []
    if set(before) != set(after) or len(after) != 149:
        errors.append('Expected exactly the same 149 components')
    for name in sorted(set(before) & set(after)):
        left, right = before[name], after[name]
        if left[0] != right[0]:
            errors.append('Component type differs: ' + name)
        if normalize(left[1]) != normalize(right[1]):
            changed.append(name)
            if name != 'modTestContourMesh':
                errors.append('Unexpected changed component: ' + name)
            else:
                if len(re.findall(BLOCK, left[1])) != 1 or len(re.findall(BLOCK, right[1])) != 1:
                    errors.append('Private native check is not unique')
                elif normalize(re.sub(BLOCK, '<NATIVE_CHECK>', left[1])) != normalize(re.sub(BLOCK, '<NATIVE_CHECK>', right[1])):
                    errors.append('Changes outside the private native check')
    callers = []
    for name, (_, body) in after.items():
        for line in body.splitlines():
            if 'TestCheckNativeFullExport' in line and not line.startswith('Private Sub '):
                callers.append({'component': name, 'line': line.strip()})
    if len(callers) != 1 or callers[0]['component'] != 'modTestContourMesh':
        errors.append('Unexpected native-check callers')
    native_entry = re.search(r'(?ms)^Public Function RunNativeContourMeshTests\(.*?^End Function\b', after['modTestContourMesh'][1])
    if not native_entry or 'TestCheckNativeFullExport doc, imported' not in native_entry.group():
        errors.append('Caller is not isolated in the native-only entrypoint')
    result = {'status': 'PASS' if not errors else 'FAIL', 'testedWorkbookSHA256': digest(args.before_book),
              'releaseWorkbookSHA256': digest(args.after_book), 'beforeDumpSHA256': digest(args.before_dump),
              'afterDumpSHA256': digest(args.after_dump), 'components': len(after), 'changed': changed,
              'changedMethod': 'modTestContourMesh.TestTestCheckNativeFullExport', 'productionUnchanged': not errors,
              'unchangedOtherMethods': not errors, 'callers': callers, 'errors': errors,
              'normalization': 'CRLF, trailing blank lines, identifier case; strings/comments preserved'}
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({k: result[k] for k in ('status', 'components', 'changed', 'errors')}))
    return bool(errors)


if __name__ == '__main__':
    raise SystemExit(main())
