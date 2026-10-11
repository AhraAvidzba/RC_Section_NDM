"""Quantify approved analytic-arc changes without widening any tolerance."""

import argparse
import json
import math
from pathlib import Path


def summarize(comparison):
    allowed = {
        "BAR": {"x", "y"},
        "LAYOUT": {"x", "y"},
        "ANCHOR": {"x1", "y1", "x2", "y2"},
        "ANNOTATION": {"x1", "y1", "x2", "y2"},
        "NDS": set("e0 kx ky n mx my min_cb max_cb min_s max_s residual_n residual_mx residual_my min_cb_strain max_cb_strain min_s_strain max_s_strain".split()),
        "FIBER_STATE": {"strain", "stress"},
        "BAR_STATE": {"strain", "stress"},
        "BATCH": {"capacity_lambda", "crack_lambda", "width", "area"},
    }
    errors, groups = [], {}
    for change in comparison["errors"]:
        if change["issue"] != "field":
            errors.append(change)
            continue
        fixture, kind, _ = change["record"]
        field = change["field"]
        if kind not in allowed or field not in allowed[kind]:
            errors.append(change)
            continue
        if kind in ("BAR", "LAYOUT", "ANCHOR", "ANNOTATION") and not fixture.endswith("case1"):
            errors.append(change)
            continue
        before, after = change["expected"], change["actual"]
        if not isinstance(before, (int, float)) or not isinstance(after, (int, float)) or not math.isfinite(before) or not math.isfinite(after):
            errors.append(change)
            continue
        key = fixture, kind, field
        metric = groups.setdefault(key, {"fixture": fixture, "kind": kind, "field": field, "changes": 0, "max_absolute": 0, "max_percent": 0})
        metric["changes"] += 1
        metric["max_absolute"] = max(metric["max_absolute"], abs(after - before))
        if before != 0:
            metric["max_percent"] = max(metric["max_percent"], 100 * abs(after - before) / abs(before))
    return {"status": "REVIEWED_ANALYTIC_CHANGE" if not errors else "FAIL", "unexpected_changes": errors, "metrics": list(groups.values())}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("comparisons", type=Path, nargs="+")
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    reports = [{"comparison": str(path), **summarize(json.loads(path.read_text(encoding="utf-8-sig")))} for path in args.comparisons]
    result = {"status": "FAIL" if any(report["unexpected_changes"] for report in reports) else "REVIEWED_ANALYTIC_CHANGE", "reports": reports}
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": result["status"], "unexpected_changes": sum(len(report["unexpected_changes"]) for report in reports)}))
    return result["status"] == "FAIL"


if __name__ == "__main__":
    raise SystemExit(main())
