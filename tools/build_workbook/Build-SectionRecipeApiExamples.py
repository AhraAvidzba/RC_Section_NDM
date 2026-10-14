"""Привязка каждого члена API к реальному VBA-вызову и native-журналу."""
from __future__ import annotations

import json
import hashlib
import re
from collections import defaultdict
from pathlib import Path

from Build_SectionRecipeApiSupport import ROOT, DOCS, inventory


def procedures(path):
    owner = re.search(r'Attribute VB_Name = "([^"]+)"', path.read_text(encoding="utf-8-sig")).group(1)
    result, current = [], None
    pending, first = [], 0
    for number, line in enumerate(path.read_text(encoding="utf-8-sig").splitlines(), 1):
        if not pending: first = number
        pending.append(line)
        if line.rstrip().endswith(" _"): continue
        raw = "\n".join(pending); pending = []
        logical = raw.replace(" _\n", " ").strip()
        match = re.match(r"(Public|Private|Friend) (?:Function|Sub|Property Get) (\w+)\(", logical)
        if match:
            current = {"owner": owner, "name": match[2], "public": match[1] == "Public", "source": path.relative_to(ROOT).as_posix(), "lines": []}
            result.append(current)
        if current is not None:
            # Строковые литералы/комментарии не считаем вызовом API.
            without_strings = re.sub(r'"(?:[^"]|"")*"', '""', logical)
            code = without_strings.split("'", 1)[0]
            current["lines"].append((first, raw, code))
            if re.match(r"End (?:Function|Sub|Property)\b", logical): current = None
    return result


def evidence():
    result = defaultdict(list)
    for folder in ("GeometryDeclarative_2026-10-11", "GeometryDeclarative_2026-10-12"):
        for path in (ROOT / "docs/regression" / folder).rglob("*.txt"):
            text = path.read_text(encoding="utf-8-sig", errors="replace")
            matches = list(re.finditer(r"(?m)^===== (\w+\.\w+) =====\s*$", text))
            for index, match in enumerate(matches):
                body = text[match.end():matches[index+1].start() if index+1 < len(matches) else len(text)]
                if re.search(r"failed=0\b", body, re.I) and not re.search(r"failed=[1-9]|SUITE_ERROR|WATCHDOG_TIMEOUT", body, re.I):
                    result[match[1]].append(path.relative_to(ROOT).as_posix())
    # Новые адресные прогоны: отдельный текст каждого entrypoint и завершённый Run.json.
    folders = [ROOT / "docs/regression/ExcelGuardGuide_2026-10-13" / name for name in
               ("SolidFinal_2026-10-14", "Audit02_2026-10-14", "Audit02Closeout_2026-10-14", "ResolveAnchor_2026-10-15")]
    for receipt in (path for folder in folders for path in folder.glob("*/Run.json")):
        run = json.loads(receipt.read_text(encoding="utf-8-sig"))
        if not run.get("Passed") or not run.get("ProductionUnchanged"):
            continue
        for item in run.get("Runs", []):
            if not item.get("Passed") or item.get("Snapshot"):
                continue
            path = receipt.parent / (item["Macro"] + ".txt")
            body = path.read_text(encoding="utf-8-sig")
            if re.search(r"(?m)^TOTAL[^\r\n]*failed=0\b", body) and not re.search(r"(?m)^FAIL|failed=[1-9]|RUNTIME ERROR", body):
                result[item["Macro"]].insert(0, path.relative_to(ROOT).as_posix())
    return result


def build():
    members = inventory()["members"]
    paths = list((ROOT / "tests").rglob("*.bas")) + list((DOCS / "examples").glob("*.bas")) + list((DOCS / "examples").glob("*.cls"))
    routines = [routine for path in paths for routine in procedures(path)]
    proofs = evidence()
    by_owner = defaultdict(list)
    for routine in routines: by_owner[routine["owner"]].append(routine)
    # Helper может вызываться через несколько уровней другого helper.
    reachable = {}
    for routine in routines:
        name = routine["owner"] + "." + routine["name"]
        if name not in proofs: continue
        found, queue = {routine["name"]}, [routine]
        while queue:
            parent = queue.pop()
            corpus = "\n".join(line[2] for line in parent["lines"])
            for child in by_owner[routine["owner"]]:
                if child["name"] not in found and re.search(r"\b" + child["name"] + r"\b", corpus, re.I):
                    found.add(child["name"]); queue.append(child)
        reachable[name] = found
    calls = defaultdict(list)
    for routine in routines:
        corpus = "\n".join(line[2] for line in routine["lines"])
        typed = defaultdict(set)
        for variable, kind in re.findall(r"\b(\w+)\s+As\s+(?:New\s+)?(\w+)\b", corpus): typed[kind].add(variable)
        for member in members:
            names = typed[member["owner"]]
            pattern = r"\b(?:" + "|".join(re.escape(name) for name in names) + r")\." + member["name"] + r"\b" if names else None
            if member["kind"] in ("Enum", "Type", "Const") or member["owner"].startswith("mod"):
                selectors = [member["name"]]
                if member["kind"] == "Enum": selectors += [field.split("=")[0].strip() for field in member["fields"]]
                pattern = r"\b(?:" + "|".join(re.escape(name) for name in selectors) + r")\b"
            if pattern:
                for number, raw, code in routine["lines"]:
                    if re.search(pattern, code, re.I):
                        roots = []
                        for root in by_owner[routine["owner"]]:
                            name = root["owner"] + "." + root["name"]
                            if name not in proofs: continue
                            if routine["name"] in reachable[name]:
                                roots.append({"entryPoint": name, "reports": proofs[name]})
                        calls[member["id"]].append({"source": routine["source"], "line": number, "routine": routine["name"],
                            "code": raw, "context": "В теле " + routine["owner"] + "." + routine["name"] + " с его подготовленными локальными объектами/входом.", "nativeRoots": roots})
    rows, missing, unproved = {}, [], []
    for member in members:
        candidates = sorted(calls[member["id"]], key=lambda item: (not bool(item["nativeRoots"]), len(item["code"])))
        if not candidates:
            missing.append(member["id"])
            continue
        selected = candidates[0]
        if not selected["nativeRoots"]: unproved.append(member["id"])
        rows[member["id"]] = selected
    source_names = sorted({example["source"] for example in rows.values()})
    report_names = sorted({report for example in rows.values() for proof in example["nativeRoots"] for report in proof["reports"]})
    status = "SOURCE_CALLS_AND_NATIVE_ROOTS_MATCH" if not missing and not unproved else "CALL_REFERENCE_DRAFT"
    return {"schema": 1, "status": status, "members": len(members), "examples": rows,
            "missingCalls": missing, "missingNativeRoots": unproved,
            "sourceHashes": {name: hashlib.sha256((ROOT/name).read_bytes()).hexdigest() for name in source_names},
            "reportHashes": {name: hashlib.sha256((ROOT/name).read_bytes()).hexdigest() for name in report_names},
            "note": "Фрагменты исполняются в явно указанном полном контексте теста/fixture; это не самостоятельные процедуры. Root подтверждён журналом; условные ветви дополнительно проверяются смысловыми assertions."}


if __name__ == "__main__":
    result = build()
    (DOCS / "SECTION_RECIPE_API_EXAMPLES.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({key: result[key] for key in ("status", "members", "missingCalls", "missingNativeRoots")}, ensure_ascii=False))
