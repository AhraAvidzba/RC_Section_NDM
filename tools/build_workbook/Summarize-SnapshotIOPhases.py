"""ДЛЯ ТЕСТОВ: сводит реальные измерения IO, проверяет равенство и освобождение контекста."""
from __future__ import annotations

import argparse
import json
import statistics
from pathlib import Path


def read_records(directory: Path) -> list[dict]:
    """Читает только законченные непрогревочные измерения, без пользовательских матриц."""
    records = json.loads((directory / "Measurements.json").read_text(encoding="utf-8-sig"))
    return [record for record in records if not record["Warmup"]]


def summarize(records: list[dict], minimum: int) -> list[dict]:
    """Считает медианы внутри одинакового сочетания режима, размера и пути показа."""
    groups: dict[tuple, list[dict]] = {}
    for record in records:
        if not record["ReadContextReleased"] or not record["ManualContextReleased"]:
            raise ValueError("Snapshot read context was not released")
        groups.setdefault((record["Count"], record["Mode"], record["Path"]), []).append(record)
    result = []
    for (count, mode, path), rows in sorted(groups.items()):
        if len(rows) < minimum:
            raise ValueError(f"Too few observations: {count}/{mode}/{path}: {len(rows)}")
        if len({row["Fingerprint"] for row in rows}) != 1:
            raise ValueError(f"Nonidentical results: {count}/{mode}/{path}")
        phases = sorted({name for row in rows for name in row["Phases"]})
        manual_phases = sorted({name for row in rows for name in row["ManualPhases"]})
        result.append({
            "Count": count, "Mode": mode, "Path": path, "Observations": len(rows),
            "Median": statistics.median(row["Seconds"] for row in rows),
            "Min": min(row["Seconds"] for row in rows),
            "Max": max(row["Seconds"] for row in rows),
            "Phases": {name: statistics.median(row["Phases"].get(name, 0) for row in rows) for name in phases},
            "ManualMedian": statistics.median(row["ManualRefreshSeconds"] for row in rows),
            "ManualPhases": {name: statistics.median(row["ManualPhases"].get(name, 0) for row in rows) for name in manual_phases},
            "SheetMedian": statistics.median(sum(row["Phases"].get(name, 0) for name in ("WriteBase", "WriteSummary", "ClearSheetExtra", "WriteFileNotice")) for row in rows),
            "Fingerprint": rows[0]["Fingerprint"], "Released": True,
        })
    return result


def main() -> None:
    """Требует одинаковые результаты между версиями и проверяет направления чтения."""
    parser = argparse.ArgumentParser()
    parser.add_argument("--baseline", type=Path)
    parser.add_argument("--candidate", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--minimum", type=int, default=5)
    args = parser.parse_args()
    candidate_records = read_records(args.candidate)
    candidate = summarize(candidate_records, args.minimum)
    baseline = summarize(read_records(args.baseline), args.minimum) if args.baseline else []
    baseline_by_case = {(row["Count"], row["Mode"], row["Path"]): row for row in baseline}
    candidate_by_case = {(row["Count"], row["Mode"], row["Path"]): row for row in candidate}
    reference_fingerprints: dict[int, str] = {}
    comparisons = []
    for row in baseline + candidate:
        reference = reference_fingerprints.setdefault(row["Count"], row["Fingerprint"])
        if reference != row["Fingerprint"]:
            raise ValueError(f"Results or plot changed between paths/versions: count={row['Count']}")
    for record in candidate_records:
        phases = record["Phases"]
        if phases.get("ReadPhysicalAfterSolve", 0):
            raise ValueError("Candidate reads Results after solving before saving the snapshot")
        if record["Path"] == "Memory" and (phases.get("ReadBinary", 0) or phases.get("LoadSheet", 0)):
            raise ValueError("Automatic memory path reloads saved data")
        if record["Mode"] == "File" and record["ManualPhases"].get("ReadBinary", 0) < 1:
            raise ValueError("Manual refresh did not load the saved snapshot")
        if record["Mode"] == "Excel" and record["ManualPhases"].get("ReadBinary", 0):
            raise ValueError("Manual Excel refresh read a binary snapshot")
    for key, after in candidate_by_case.items():
        before = baseline_by_case.get(key)
        if before:
            comparisons.append({
                "Count": key[0], "Mode": key[1], "Path": key[2],
                "Before": before["Median"], "After": after["Median"],
                "ReductionPercent": 100 * (before["Median"] - after["Median"]) / before["Median"],
                "ManualBefore": before["ManualMedian"], "ManualAfter": after["ManualMedian"],
            })
    memory_comparisons = []
    for (count, mode, path), memory in candidate_by_case.items():
        if path != "Memory":
            continue
        reload = candidate_by_case.get((count, mode, "Reload"))
        if reload:
            memory_comparisons.append({"Count": count, "Mode": mode, "Memory": memory["Median"], "Reload": reload["Median"],
                                       "MemoryReductionPercent": 100 * (reload["Median"] - memory["Median"]) / reload["Median"]})
    proof = {"Status": "PASS", "Baseline": baseline, "Candidate": candidate, "Comparisons": comparisons,
             "MemoryComparisons": memory_comparisons, "ExactResultAndPlotFingerprints": True,
             "AutomaticSnapshotReloads": 0, "PostSolveSheetReadsBeforeSave": 0, "ManualRefreshLoadsSavedData": True,
             "AllReadContextsReleased": True, "NumberEquivalence": "Only IEEE-754 positive/negative exact zero; all other values exact"}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(proof, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"Status": proof["Status"], "Cases": len(candidate), "Comparisons": comparisons, "MemoryComparisons": memory_comparisons}))


if __name__ == "__main__":
    main()
