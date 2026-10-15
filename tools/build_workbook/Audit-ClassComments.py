"""Проверяет русские шапки классов и комментарии непосредственно перед методами."""
from __future__ import annotations
import argparse
import json
import re
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
METHOD = re.compile(r"^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)", re.I)
RUSSIAN = re.compile(r"[А-Яа-яЁё]")

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
                            "commented": own_comment and any(RUSSIAN.search(s) for s in comments), "comments": comments})
    return {"path": path.relative_to(ROOT).as_posix(), "header": header,
            "headerRussian": any(RUSSIAN.search(s) for s in header),
            "headerStyle": len(header) >= 3 and header[0] == "' " + '='*74
                and header[2] == header[0] and bool(RUSSIAN.search(header[1])),
            "methods": methods}

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
    args = parser.parse_args()
    files = [p for folder in ("src", "tests", "docs/geometry_refactor/examples") for p in (ROOT/folder).rglob("*.cls")]
    classes = [inspect(p) for p in sorted(files)]
    missing = [{"path": c["path"], **m} for c in classes for m in c["methods"] if not m["commented"]]
    headers = [c["path"] for c in classes if not c["headerRussian"]]
    styles = [c["path"] for c in classes if not c["headerStyle"]]
    code_changes = []
    changed_comments = []
    if args.base_ref:
        for path, current in zip(sorted(files), classes):
            old = subprocess.check_output(['git', 'show', args.base_ref+':'+current['path']], cwd=ROOT).decode('utf-8-sig')
            new = path.read_text(encoding='utf-8-sig')
            if executable_lines(old) != executable_lines(new): code_changes.append(current['path'])
            before = inspect(path, old)
            old_methods = {(m['kind'].lower(),m['name'].lower()):m['comments'] for m in before['methods']}
            edited = [m['name'] for m in current['methods']
                      if m['comments'] != old_methods.get((m['kind'].lower(),m['name'].lower()))]
            if old != new:
                changed_comments.append({'path':current['path'], 'headerChanged':before['header'] != current['header'],
                                         'methodCount':len(edited), 'methods':edited})
    failed = missing or headers or styles or code_changes
    result = {"schema": 2, "status": "PASS" if not failed else "REVIEW_REQUIRED",
              "classes": len(classes), "methods": sum(len(c["methods"]) for c in classes),
              "missingHeaderCount": len(headers), "missingMethodCount": len(missing),
              "missingHeaders": headers, "missingMethods": missing, "nonstandardHeaders": styles,
              "baseRef":args.base_ref, "executableCodeChanges":code_changes,
              "commentChanges":changed_comments, "details": classes}
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2)+"\n", encoding="utf-8")
    print(json.dumps({k: result[k] for k in ("status", "classes", "methods", "missingHeaderCount", "missingMethodCount", "nonstandardHeaders", "executableCodeChanges")}))
    if args.check and failed: raise SystemExit(1)

if __name__ == "__main__":
    main()
