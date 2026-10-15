"""Проверяет шапки, пояснения методов и полей, а также расположение тестовых методов."""
from __future__ import annotations
import argparse
import json
import re
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
METHOD = re.compile(r"^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)", re.I)
RUSSIAN = re.compile(r"[А-Яа-яЁё]")
TEST_LABEL = re.compile(r"^\s*'\s*ДЛЯ ТЕСТОВ(?::|\s*$)", re.I)

def method_spans(lines):
    """Выделяет процедуры целиком, чтобы сравнение допускало только их перестановку."""
    spans = []
    for index, line in enumerate(lines):
        match = METHOD.match(line)
        if not match: continue
        end_kind = match[2].split()[0]
        end = next(i+1 for i in range(index, len(lines))
                   if re.search(r"(?:^|:)\s*End\s+"+end_kind+r"\s*$",
                                '\n'.join(executable_lines(lines[i])), re.I))
        spans.append((index, end, match[2].lower(), match[3].lower()))
    return spans

def code_units(text):
    """Точный исполняемый текст деклараций и каждой процедуры, без комментариев."""
    lines = text.splitlines()
    declarations = []
    methods = {}
    previous = 0
    for start, end, kind, name in method_spans(lines):
        declarations.extend(lines[previous:start])
        key = (kind, name)
        if key in methods: raise ValueError('Duplicate procedure: '+str(key))
        methods[key] = executable_lines('\n'.join(lines[start:end]))
        previous = end
    declarations.extend(lines[previous:])
    return executable_lines('\n'.join(declarations)), methods

def fields_inspection(lines):
    """Проверяет декларации класса и поля Type, не затрагивая локальные переменные."""
    spans = method_spans(lines)
    inside_type = False
    fields = []
    for index, line in enumerate(lines):
        if any(start <= index < end for start, end, _, _ in spans): continue
        if re.match(r'\s*(?:Public|Private)\s+Type\b', line, re.I):
            inside_type = True
            continue
        if re.match(r'\s*End\s+Type\b', line, re.I):
            inside_type = False
            continue
        if not line.strip() or line.lstrip().startswith("'"): continue
        declaration = inside_type or re.match(
            r'\s*(?:Public|Private|Friend|Dim)\s+(?!Type\b|Enum\b|Declare\b|Event\b)', line, re.I)
        if not declaration: continue
        code = executable_lines(line)
        prefix = code[0] if code else ''
        comment = line[len(prefix):].lstrip()
        fields.append({'line':index+1, 'declaration':prefix.strip(),
                       'comment':comment, 'commented':comment.startswith("'") and bool(RUSSIAN.search(comment))})
    return fields

def module_test_order(path):
    """В обычных рабочих модулях тестовые процедуры также должны завершать файл."""
    if 'tests' in path.parts or 'examples' in path.parts: return []
    lines = path.read_text(encoding='utf-8-sig').splitlines()
    test_started = False
    wrong = []
    for index, line in enumerate(lines):
        match = METHOD.match(line)
        if not match: continue
        is_test = any(TEST_LABEL.search(s) for s in preceding(lines,index))
        if is_test: test_started = True
        elif test_started: wrong.append(match[3])
    return wrong

def preceding(lines, index):
    result = []
    for line in reversed(lines[:index]):
        if not line.strip():
            if result: break
            continue
        if line.lstrip().startswith("'"):
            result.append(line)
        else: break
    return list(reversed(result))

def inspect(path, text=None):
    lines = (path.read_text(encoding="utf-8-sig") if text is None else text).splitlines()
    option = next(i for i, line in enumerate(lines) if line.strip() == "Option Explicit")
    header = []
    start = option + 1
    while start < len(lines) and not lines[start].strip(): start += 1
    end = start
    while end < len(lines) and lines[end].lstrip().startswith("'"):
        header.append(lines[end])
        end += 1
    methods = []
    for index, line in enumerate(lines):
        match = METHOD.match(line)
        if match:
            comments = preceding(lines, index)
            own_comment = comments != header
            methods.append({"line": index+1, "name": match[3], "kind": match[2],
                            "commented": own_comment and any(RUSSIAN.search(s) for s in comments), "comments": comments,
                            "testOnly":own_comment and any(TEST_LABEL.search(s) for s in comments)})
    all_test = path.parts[-2] == 'tests' or path.stem.startswith(('CFake','CTest','CLinearConcrete','CLinearSteel'))
    test_started = False
    wrong_order = []
    for method in methods:
        if method['testOnly']: test_started = True
        elif test_started and not all_test: wrong_order.append(method['name'])
    return {"path": path.relative_to(ROOT).as_posix(), "header": header,
            "headerRussian": any(RUSSIAN.search(s) for s in header),
            "headerStyle": len(header) >= 3 and header[0] == "' " + '='*74
                and header[2] == header[0] and bool(RUSSIAN.search(header[1])),
            "methods": methods, "fields":fields_inspection(lines), "regularMethodsAfterTests":wrong_order}

def executable_lines(text):
    """Убирает комментарии вне строк, сохраняя точный код и строковые литералы VBA."""
    result = []
    for line in text.splitlines():
        quoted = False
        i = 0
        while i < len(line):
            if line[i] == '"':
                if quoted and i+1 < len(line) and line[i+1] == '"':
                    i += 2
                    continue
                quoted = not quoted
            elif line[i] == "'" and not quoted:
                line = line[:i]
                break
            i += 1
        line = line.rstrip()
        if line: result.append(line)
    return result

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--base-ref", help="Сравнить исполняемый VBA с указанным Git-коммитом")
    parser.add_argument("--allow-method-reorder", action="store_true",
                        help="Разрешить перестановку целых процедур при неизменных декларациях и телах")
    args = parser.parse_args()
    files = [p for folder in ("src", "tests", "docs/geometry_refactor/examples") for p in (ROOT/folder).rglob("*.cls")]
    modules = sorted(p for folder in ('src','tests','docs/geometry_refactor/examples') for p in (ROOT/folder).rglob('*.bas'))
    classes = [inspect(p) for p in sorted(files)]
    missing = [{"path": c["path"], **m} for c in classes for m in c["methods"] if not m["commented"]]
    headers = [c["path"] for c in classes if not c["headerRussian"]]
    styles = [c["path"] for c in classes if not c["headerStyle"]]
    missing_fields = [{"path":c['path'], **f} for c in classes for f in c['fields'] if not f['commented']]
    wrong_order = [{"path":c['path'], "methods":c['regularMethodsAfterTests']} for c in classes if c['regularMethodsAfterTests']]
    for path in modules:
        methods = module_test_order(path)
        if methods: wrong_order.append({'path':path.relative_to(ROOT).as_posix(),'methods':methods})
    code_changes = []
    changed_comments = []
    if args.base_ref:
        for path, current in zip(sorted(files), classes):
            old = subprocess.check_output(['git', 'show', args.base_ref+':'+current['path']], cwd=ROOT).decode('utf-8-sig')
            new = path.read_text(encoding='utf-8-sig')
            equal = code_units(old) == code_units(new) if args.allow_method_reorder else executable_lines(old) == executable_lines(new)
            if not equal: code_changes.append(current['path'])
            before = inspect(path, old)
            old_methods = {(m['kind'].lower(),m['name'].lower()):m['comments'] for m in before['methods']}
            edited = [m['name'] for m in current['methods']
                      if m['comments'] != old_methods.get((m['kind'].lower(),m['name'].lower()))]
            if old != new:
                changed_comments.append({'path':current['path'], 'headerChanged':before['header'] != current['header'],
                                         'methodCount':len(edited), 'methods':edited})
        for path in modules:
            relative = path.relative_to(ROOT).as_posix()
            old = subprocess.check_output(['git','show',args.base_ref+':'+relative],cwd=ROOT).decode('utf-8-sig')
            new = path.read_text(encoding='utf-8-sig')
            equal = code_units(old) == code_units(new) if args.allow_method_reorder else executable_lines(old) == executable_lines(new)
            if not equal: code_changes.append(relative)
    failed = missing or headers or styles or code_changes or missing_fields or wrong_order
    result = {"schema": 3, "status": "PASS" if not failed else "REVIEW_REQUIRED",
              "classes": len(classes), "methods": sum(len(c["methods"]) for c in classes),
              "sourceComponents":len(classes)+len(modules),
              "missingHeaderCount": len(headers), "missingMethodCount": len(missing),
              "missingHeaders": headers, "missingMethods": missing, "nonstandardHeaders": styles,
              "fields":sum(len(c['fields']) for c in classes), "missingFieldCount":len(missing_fields),
              "missingFields":missing_fields, "wrongTestMethodOrder":wrong_order,
              "allowMethodReorder":args.allow_method_reorder,
              "baseRef":args.base_ref, "executableCodeChanges":code_changes,
              "commentChanges":changed_comments, "details": classes}
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2)+"\n", encoding="utf-8")
    print(json.dumps({k: result[k] for k in ("status", "classes", "methods", "fields", "missingHeaderCount", "missingMethodCount", "missingFieldCount", "nonstandardHeaders", "wrongTestMethodOrder", "executableCodeChanges")}))
    if args.check and failed: raise SystemExit(1)

if __name__ == "__main__":
    main()
