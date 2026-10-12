"""Проверяет русские шапки классов и комментарии непосредственно перед методами."""
from __future__ import annotations
import argparse
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
METHOD = re.compile(r"^\s*(Public|Private|Friend)\s+(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)", re.I)
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

def inspect(path):
    lines = path.read_text(encoding="utf-8-sig").splitlines()
    option = next(i for i, line in enumerate(lines) if line.strip() == "Option Explicit")
    header = []
    for line in lines[option+1:]:
        if not line.strip(): continue
        if line.lstrip().startswith("'"): header.append(line)
        else: break
    methods = []
    for index, line in enumerate(lines):
        match = METHOD.match(line)
        if match:
            comments = preceding(lines, index)
            methods.append({"line": index+1, "name": match[3], "kind": match[2],
                            "commented": any(RUSSIAN.search(s) for s in comments), "comments": comments})
    return {"path": path.relative_to(ROOT).as_posix(), "header": header,
            "headerRussian": any(RUSSIAN.search(s) for s in header), "methods": methods}

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    files = [p for folder in ("src", "tests", "docs/geometry_refactor/examples") for p in (ROOT/folder).rglob("*.cls")]
    classes = [inspect(p) for p in sorted(files)]
    missing = [{"path": c["path"], **m} for c in classes for m in c["methods"] if not m["commented"]]
    headers = [c["path"] for c in classes if not c["headerRussian"]]
    result = {"schema": 1, "status": "PASS" if not missing and not headers else "MISSING_COMMENTS",
              "classes": len(classes), "methods": sum(len(c["methods"]) for c in classes),
              "missingHeaderCount": len(headers), "missingMethodCount": len(missing),
              "missingHeaders": headers, "missingMethods": missing, "details": classes}
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2)+"\n", encoding="utf-8")
    print(json.dumps({k: result[k] for k in ("status", "classes", "methods", "missingHeaderCount", "missingMethodCount")}))
    if args.check and (missing or headers): raise SystemExit(1)

if __name__ == "__main__":
    main()
