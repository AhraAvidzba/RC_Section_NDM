"""Сопоставляет фактические отчёты A/B без подмены ожиданий расчёта.

Существенная разница времени здесь только повод повторить замер. Проверки
физики и равновесия выполняет VBA; сравнение численных записей строгое.
"""

import argparse
import csv
import hashlib
import json
import re
import statistics
from pathlib import Path


def records(lines, prefix):
    return [line for line in lines if line.startswith(prefix)]


def fields(line):
    return dict(part.split("=", 1) for part in line.split("|")[1:] if "=" in part)


def read_report(root, record):
    path = root / record["Report"]
    raw = path.read_bytes()
    if hashlib.sha256(raw).hexdigest().upper() != record["ReportSHA256"]:
        raise ValueError(f"Report changed: {path}")
    lines = raw.decode("utf-8-sig").splitlines()
    totals = records(lines, "TOTAL_")
    if len(totals) != 1:
        raise ValueError(f"Missing or duplicate suite total: {path}")
    match = re.search(r"passed=(\d+); failed=(\d+)", totals[0])
    if not match:
        raise ValueError(f"Malformed suite total: {path}")
    timings = [fields(line) for line in records(lines, "FOCUSED_TIME|")]
    batch_times = [fields(line) for line in records(lines, "CORE_BATCH_TIME|")]
    compare = records(lines, "FOCUSED_COMPARE|")
    compare_ids = [fields(line)["id"] for line in compare]
    timing_ids = [time["id"] for time in timings]
    if len(set(compare_ids)) != len(compare_ids) or len(set(timing_ids)) != len(timing_ids):
        raise ValueError(f"Duplicate case records: {path}")
    expected = record["ExpectedCases"]
    is_batch = record["Group"] == "Batch15"
    if is_batch:
        coverage = len(compare) == expected and len(batch_times) == 1
    else:
        # Нулевая нагрузка проверяется как отклонённый ввод, результата нет.
        coverage = len(timings) == expected and len(compare) == expected - sum(
            time["id"] == "zero" for time in timings
        )
    return {
        "lines": lines,
        "timings": timings,
        "batch_times": batch_times,
        "passed": int(match[1]),
        "failed": int(match[2]),
        "coverage": coverage,
        "reopen": any("RESULTS_SAVE_REOPEN:" in x and "equal=True" in x for x in lines),
        "styles": any("STATUS_STYLE_SAVE_REOPEN:" in x and "equal=True" in x for x in lines),
        "source": "SOURCE_UNCHANGED: True" in lines,
        "unresolved": records(lines, "FOCUSED_UNRESOLVED|"),
        "failures": [x for x in lines if re.match(r"^(FAIL:|RUNTIME ERROR:|WATCHDOG_FAILURE:|RESTORE ERROR:)", x)],
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--directory", default="docs/regression/Performance/CoreRegression_2026-10-15")
    parser.add_argument("--allow-partial", action="store_true")
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[2]
    directory = root / args.directory
    progress = json.loads((directory / "Progress.json").read_text(encoding="utf-8-sig"))
    grouped = {}
    checked = []
    unresolved = []
    failures = []
    for record in progress:
        report = read_report(root, record)
        key = (record["Shape"], record["Group"], record["Mode"])
        grouped.setdefault(key, {})[record["Version"]] = report
        checked.append({
            "id": record["ID"], "passed": report["passed"], "failed": report["failed"],
            "complete": report["coverage"], "resultsReopen": report["reopen"],
            "stylesReopen": report["styles"], "sourceUnchanged": report["source"],
            "runnerPassed": record["Passed"] and record["ExitCode"] == 0,
        })
        unresolved.extend({"report": record["ID"], "line": line} for line in report["unresolved"])
        failures.extend({"report": record["ID"], "line": line} for line in report["failures"])
    rows, comparisons, differences = [], [], []
    for key, pair in grouped.items():
        if set(pair) != {"Before", "After"}:
            continue
        before, after = pair["Before"], pair["After"]
        comparison = {"shape": key[0], "group": key[1], "mode": key[2]}
        # META включает пользовательские причины, applies/calculated, имена и коды.
        for prefix in ("FOCUSED_COMPARE|", "FOCUSED_STATE|", "FOCUSED_RESULT|", "META:"):
            left = records(before["lines"], prefix)
            right = records(after["lines"], prefix)
            comparison[prefix] = {"countBefore": len(left), "countAfter": len(right), "exact": left == right}
            if left != right:
                differences.append({"key": key, "prefix": prefix,
                                    "before": left, "after": right})
        comparisons.append(comparison)
        bt = {time["id"]: time for time in before["timings"]}
        at = {time["id"]: time for time in after["timings"]}
        if set(bt) != set(at):
            raise ValueError(f"Timing coverage differs: {key}")
        for load_id in bt:
            old, new = float(bt[load_id]["elapsedSec"]), float(at[load_id]["elapsedSec"])
            delta = new - old
            rows.append({
                "shape": key[0], "group": key[1], "mode": key[2], "id": load_id,
                "beforeSeconds": old, "afterSeconds": new, "deltaSeconds": delta,
                "ratio": new / old if old else None,
                "beforeSolverCalls": int(bt[load_id]["solverCalls"]),
                "afterSolverCalls": int(at[load_id]["solverCalls"]),
                # Порог отбора повторов, не нормативный критерий приёмки.
                "repeatNeeded": delta > max(0.25, old * 0.15),
            })
        if key[1] == "Batch15":
            old = float(before["batch_times"][0]["seconds"])
            new = float(after["batch_times"][0]["seconds"])
            rows.append({"shape": key[0], "group": key[1], "mode": key[2], "id": "batch15",
                         "beforeSeconds": old, "afterSeconds": new, "deltaSeconds": new - old,
                         "ratio": new / old, "beforeSolverCalls": None, "afterSolverCalls": None,
                         "repeatNeeded": new - old > max(0.25, old * 0.15)})
    individual = [row for row in rows if row["group"] != "Batch15"]
    complete = len(checked) == 28 and len(individual) == 122 and len(rows) == 124
    summary = {
        "complete": complete, "reports": len(checked), "expectedReports": 28,
        "pairedIndividualCases": len(individual), "expectedIndividualCases": 122,
        "passedAssertionsBothVersions": sum(report["passed"] for report in checked),
        "failedAssertionsBothVersions": sum(report["failed"] for report in checked),
        "allReportChecksPassed": all(report["complete"] and not report["failed"] and
                                     report["resultsReopen"] and report["stylesReopen"] and
                                     report["sourceUnchanged"] and report["runnerPassed"] for report in checked),
        "allPairedOutputsExact": not differences, "differentOutputGroups": len(differences),
        "maxBeforeSeconds": max((row["beforeSeconds"] for row in individual), default=None),
        "maxAfterSeconds": max((row["afterSeconds"] for row in individual), default=None),
        "medianTimeDeltaSeconds": statistics.median(row["deltaSeconds"] for row in individual) if individual else None,
        "timingRepeatCandidates": [row for row in rows if row["repeatNeeded"]],
        "unresolved": unresolved, "failures": failures, "reportsChecked": checked,
        "comparison": comparisons,
    }
    (directory / "Comparison.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (directory / "OutputDifferences.json").write_text(json.dumps(differences, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    with (directory / "TimingComparison.csv").open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=["shape", "group", "mode", "id", "beforeSeconds", "afterSeconds", "deltaSeconds", "ratio", "beforeSolverCalls", "afterSolverCalls", "repeatNeeded"])
        writer.writeheader()
        writer.writerows(rows)
    print(json.dumps({key: value for key, value in summary.items() if key not in ("reportsChecked", "comparison", "unresolved")}, ensure_ascii=False, indent=2))
    print(f"Unresolved entries: {len(unresolved)}")
    if not args.allow_partial and not complete:
        raise SystemExit("Incomplete evidence; no acceptance decision is possible.")
    if failures or summary["failedAssertionsBothVersions"] or not summary["allReportChecksPassed"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
