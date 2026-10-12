"""Инвентаризация сохранённых геометрических доказательств без запуска Excel."""
from __future__ import annotations

import argparse
import collections
import hashlib
import json
import re
from pathlib import Path


def inventory(root: Path, folder: Path) -> list[dict]:
    """Читаем каждый журнал; исторический FAIL не превращается в текущий PASS."""
    entries = []
    for path in sorted(folder.rglob("*")):
        if not path.is_file() or path.suffix.lower() not in {".txt", ".log", ".json", ".md"}:
            continue
        raw = path.read_bytes()
        text = raw.decode("utf-8-sig", errors="replace")
        totals = re.findall(r"(?m)^TOTAL[^\r\n]*", text)
        failures = re.findall(r"(?m)^(?:FAIL:|RUNTIME ERROR:|WATCHDOG_FAILURE:)[^\r\n]*", text)
        status = None
        if path.suffix.lower() == ".json":
            try:
                obj = json.loads(text)
                if isinstance(obj, dict):
                    status = obj.get("Status", obj.get("status"))
            except ValueError:
                status = "INVALID_JSON"
        entries.append({
            "path": path.relative_to(root).as_posix(), "bytes": len(raw),
            "sha256": hashlib.sha256(raw).hexdigest(), "declaredStatus": status,
            "totals": totals, "failureCount": len(failures),
            "failureSamples": failures[:4],
            "watchdog": re.findall(r"(?m)^WATCHDOG_COMPLETED:[^\r\n]*", text),
            "historicalProbe": bool(re.search(r"Failed|Rejected|Probe", path.as_posix())),
        })
    return entries


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--folder", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[2]
    entries = inventory(root, root / args.folder)
    target = root / args.output
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps({"schema": 1, "scope": args.folder, "files": entries},
                                 ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"files": len(entries), "declaredStatuses": dict(collections.Counter(
        str(e["declaredStatus"]) for e in entries if e["declaredStatus"])),
        "filesWithFailures": sum(e["failureCount"] > 0 for e in entries),
        "historicalProbes": sum(e["historicalProbe"] for e in entries)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
