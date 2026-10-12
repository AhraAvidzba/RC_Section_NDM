"""Audit fixture snapshots with the existing comparator and deliberate mutations."""

import argparse
import hashlib
import json
import runpy
import struct
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("baseline", type=Path)
    parser.add_argument("actual", type=Path)
    parser.add_argument("--report", required=True, type=Path)
    args = parser.parse_args()
    api = runpy.run_path(str(Path(__file__).with_name("Compare-GeometrySnapshots.py")))
    expected = api["read_snapshot"](args.baseline)
    actual = api["read_snapshot"](args.actual)
    errors = api["compare"](expected, actual)
    exact = api["compare"](expected, actual, exact=True)
    metrics = {}
    for error in exact:
        if error["issue"] == "field" and isinstance(error["expected"], float) and isinstance(error["actual"], float):
            key = ".".join(error["record"][:2] + [error["field"]])
            metrics[key] = max(metrics.get(key, 0.0), abs(error["actual"] - error["expected"]))
    mutations = []
    # These numeric fixtures have no contour records; contour roles have a
    # separate native ownership/round-trip suite. Do not fabricate such rows.
    for kind, field in (("CONCRETE", "x"), ("BAR", "diameter"), ("FIBER", "fill")):
        changed = [row[:] for row in expected]
        index = next(i for i, row in enumerate(changed) if row[1] == kind)
        column = 3 + api["FIELDS"][kind].index(field)
        changed[index][column] = "d:" + struct.pack(">d", api["decode"](changed[index][column]) + 1.0).hex().upper()
        failures = api["compare"](expected, changed)
        detected = len(failures) == 1 and failures[0].get("field") == field
        mutations.append({"mutation": f"{kind}.{field}", "detected": detected})
    status = "PASS" if not errors and all(item["detected"] for item in mutations) else "FAIL"
    report = {
        "status": status,
        "baseline": str(args.baseline),
        "baseline_sha256": hashlib.sha256(args.baseline.read_bytes()).hexdigest().upper(),
        "actual": str(args.actual),
        "actual_sha256": hashlib.sha256(args.actual.read_bytes()).hexdigest().upper(),
        "records": len(actual),
        "errors": errors,
        "exact_difference_count": len(exact),
        "exact_differences_by_field": metrics,
        "mutations": mutations,
        "tolerances": "Unchanged Compare-GeometrySnapshots.py",
    }
    args.report.write_text(json.dumps(report, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps({"status": status, "records": len(actual), "errors": len(errors), "mutations": len(mutations)}))
    return 0 if status == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
