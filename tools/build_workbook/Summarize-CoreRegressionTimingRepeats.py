"""Проверяет три повторных A/B замера и неизменность результатов."""

import argparse
import importlib.util
import json
import statistics
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location("core_summary", Path(__file__).with_name("Summarize-CoreRegression.py"))
CORE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(CORE)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--directory", default="docs/regression/Performance/CoreRegression_2026-10-15")
    parser.add_argument("--allow-partial", action="store_true")
    args = parser.parse_args()
    directory = ROOT / args.directory
    original = json.loads((directory / "Progress.json").read_text(encoding="utf-8-sig"))
    progress_path = directory / "TimingRepeatsProgress.json"
    progress = json.loads(progress_path.read_text(encoding="utf-8-sig")) if progress_path.exists() else []
    summary = json.loads((directory / "Comparison.json").read_text(encoding="utf-8-sig"))
    if not summary["complete"]:
        raise ValueError("Complete main comparison required before timing repeats.")
    groups = {(row["shape"], row["group"], row["mode"]) for row in summary["timingRepeatCandidates"]}
    bases = {(record["Version"], record["Shape"], record["Group"], record["Mode"]): record for record in original}
    reports, timings = [], {}
    for record in progress:
        report = CORE.read_report(ROOT, record)
        key = (record["Version"], record["Shape"], record["Group"], record["Mode"])
        base = CORE.read_report(ROOT, bases[key])
        outputs_equal = all(CORE.records(report["lines"], prefix) == CORE.records(base["lines"], prefix)
                            for prefix in ("FOCUSED_COMPARE|", "FOCUSED_STATE|", "FOCUSED_RESULT|", "META:"))
        reports.append({"id": record["ID"], "passed": report["passed"], "failed": report["failed"],
                        "valid": record["Passed"] and record["ExitCode"] == 0 and report["coverage"] and
                                 report["reopen"] and report["styles"] and report["source"] and
                                 not report["failures"] and not report["failed"],
                        "outputsExact": outputs_equal, "repeat": record["Repeat"]})
        times = report["timings"]
        if record["Group"] == "Batch15":
            times = [{"id": "batch15", "elapsedSec": report["batch_times"][0]["seconds"]}]
        for time in times:
            time_key = (record["Shape"], record["Group"], record["Mode"], time["id"])
            timings.setdefault(time_key, {}).setdefault(record["Version"], []).append({
                "repeat": record["Repeat"], "seconds": float(time["elapsedSec"])})
    rows, incomplete = [], []
    for key, pair in timings.items():
        if set(pair) != {"Before", "After"} or any(sorted(item["repeat"] for item in times) != [1, 2, 3] for times in pair.values()):
            incomplete.append(key)
            continue
        before = [item["seconds"] for item in sorted(pair["Before"], key=lambda item: item["repeat"])]
        after = [item["seconds"] for item in sorted(pair["After"], key=lambda item: item["repeat"])]
        old, new = statistics.median(before), statistics.median(after)
        delta = new - old
        paired_delta = statistics.median(new_time - old_time for old_time, new_time in zip(before, after))
        rows.append({"shape": key[0], "group": key[1], "mode": key[2], "id": key[3],
                     "beforeSeconds": before, "afterSeconds": after, "beforeMedianSeconds": old,
                     "afterMedianSeconds": new, "deltaMedianSeconds": delta, "medianRatio": new / old if old else None,
                     "medianPairedDeltaSeconds": paired_delta,
                     "persistentSlowdown": max(delta, paired_delta) > max(0.25, old * 0.15)})
    complete = not incomplete and len(reports) == len(groups) * 6
    output = {
        "Complete": complete, "Reports": len(reports), "ExpectedReports": len(groups) * 6,
        "AllAssertionsPassed": all(report["valid"] for report in reports),
        "AllOutputsExact": all(report["outputsExact"] for report in reports),
        "PersistentSlowdowns": [row for row in rows if row["persistentSlowdown"]],
        "PassedAssertionsBothVersions": sum(report["passed"] for report in reports),
        "FailedAssertionsBothVersions": sum(report["failed"] for report in reports),
        "IncompleteCases": incomplete, "Records": rows, "ReportsChecked": reports,
    }
    (directory / "TimingRepeatComparison.json").write_text(json.dumps(output, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({key: output[key] for key in ("Complete", "Reports", "ExpectedReports", "AllAssertionsPassed", "AllOutputsExact", "PersistentSlowdowns")}, ensure_ascii=False, indent=2))
    if (not args.allow_partial and not complete) or not output["AllAssertionsPassed"] or not output["AllOutputsExact"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
