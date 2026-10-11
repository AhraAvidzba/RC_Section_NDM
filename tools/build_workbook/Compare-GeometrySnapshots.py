"""Compare native VBA snapshot format 2 without losing Double precision."""

import argparse
import csv
import json
import math
import struct
from pathlib import Path


FIELDS = {
    "MODEL": "source mesh_source concrete_count rebar_count contour_count annotation_count".split(),
    "PROPS": "area sx sy cx cy ix iy ixy ixc iyc ixyc i1 i2 angle rx ry r1 r2".split(),
    "BOUNDS": "min_x max_x min_y max_y".split(),
    "CONCRETE": "x y area material shape width height rotation ix iy ixy source handle comment".split(),
    "CONTOUR": "type loop role x1 y1 x2 y2 cx cy radius sweep source comment".split(),
    "BAR": "x y diameter area material steel source handle comment".split(),
    "ANNOTATION": "type x1 y1 x2 y2 nx ny text value unit comment".split(),
    "FIBER": "x y area fill width height material".split(),
    "LAYOUT": "x y diameter area steel group comment".split(),
    "ANCHOR": "x1 y1 x2 y2 nx ny distance".split(),
    "INPUT": "value formula".split(),
    "NDS": "converged e0 kx ky n mx my min_cb max_cb min_s max_s residual_n residual_mx residual_my min_cb_strain max_cb_strain min_s_strain max_s_strain".split(),
    "LOAD": "n mx my ref_x ref_y".split(),
    "FIBER_STATE": "strain stress".split(),
    "BAR_STATE": "strain stress".split(),
    "PLOT": "left top width height plot_left plot_top plot_width plot_height series shapes".split(),
    "PLOT_POINT": "series x y".split(),
    "PLOT_SHAPE": "type left top width height rotation text".split(),
    "BATCH": "status crack_status capacity_status formation_status width_calculated capacity_lambda crack_lambda width area".split(),
}


def decode(token):
    tag, content = token.split(":", 1)
    if tag == "d":
        if len(content) != 16:
            raise ValueError(f"Invalid Double: {token}")
        return struct.unpack(">d", bytes.fromhex(content))[0]
    if tag == "i":
        return int(content)
    if tag == "b":
        if content not in ("0", "1"):
            raise ValueError(f"Invalid Boolean: {token}")
        return content == "1"
    if tag == "s":
        return content.replace("%09", "\t").replace("%0D", "\r").replace("%0A", "\n").replace("%25", "%")
    if tag in ("e", "n"):
        return None
    raise ValueError(f"Unknown field encoding: {token}")


def read_snapshot(path):
    records = []
    with path.open(encoding="utf-8-sig", newline="") as stream:
        for row in csv.reader(stream, delimiter="\t"):
            if row and row[0].startswith("FAIL:"):
                raise ValueError(f"Native snapshot failed: {row[0]}")
            if len(row) < 3 or row[1] not in FIELDS:
                continue
            if len(row) != 3 + len(FIELDS[row[1]]):
                raise ValueError(f"Unexpected schema in {row[:3]}: {len(row)} columns")
            for token in row[3:]:
                decode(token)
            records.append(row)
    if not records:
        raise ValueError("No snapshot records found")
    return records


def tolerance(kind, name):
    if kind == "NDS":
        if name == "e0" or name.endswith("_strain"):
            return 1e-10, 1e-6
        if name in ("kx", "ky"):
            return 1e-13, 1e-6
        if name in ("n", "residual_n"):
            return 0.01, 1e-6
        if name in ("mx", "my", "residual_mx", "residual_my"):
            return 1.0, 1e-6
        return 1e-6, 1e-6
    if kind in ("FIBER_STATE", "BAR_STATE"):
        return (1e-10 if name == "strain" else 1e-6), 1e-6
    if kind == "BATCH":
        if name in ("capacity_lambda", "crack_lambda"):
            return 1e-8, 1e-6
        return (1e-6 if name == "area" else 1e-7), 1e-6
    if name in ("area",):
        return 1e-6, 1e-12
    if name in ("sx", "sy"):
        return 1e-4, 1e-12
    if name in ("ix", "iy", "ixy", "ixc", "iyc", "ixyc", "i1", "i2"):
        return 1e-3, 1e-12
    if name in ("angle", "sweep", "rotation", "nx", "ny", "fill"):
        return 1e-12, 1e-12
    return 1e-7, 1e-12


def compare(expected, actual, exact=False, skip_input=False):
    if skip_input:
        expected = [row for row in expected if row[1] != "INPUT"]
        actual = [row for row in actual if row[1] != "INPUT"]
    errors = []
    if len(expected) != len(actual):
        errors.append({"issue": "record_count", "expected": len(expected), "actual": len(actual)})
    for index, (left, right) in enumerate(zip(expected, actual)):
        if left[:3] != right[:3]:
            errors.append({"issue": "order_or_id", "index": index, "expected": left[:3], "actual": right[:3]})
            continue
        for name, a, b in zip(FIELDS[left[1]], left[3:], right[3:]):
            av, bv = decode(a), decode(b)
            equal = a == b
            if isinstance(av, float) and not math.isfinite(av):
                equal = False
            if isinstance(bv, float) and not math.isfinite(bv):
                equal = False
            if not exact and a.startswith("d:") and b.startswith("d:") and math.isfinite(av) and math.isfinite(bv):
                absolute, relative = tolerance(left[1], name)
                equal = abs(av - bv) <= absolute + relative * max(abs(av), abs(bv))
            if not equal:
                errors.append({"issue": "field", "record": left[:3], "field": name, "expected": av, "actual": bv})
    return errors


def map_generated_rebar_names(expected, actual):
    """Record one-to-one provenance names; never alias contours or model IDs."""
    column = 3 + FIELDS["BAR"].index("source")
    before = {tuple(row[:3]): row for row in expected if row[1] == "BAR"}
    after = {tuple(row[:3]): row for row in actual if row[1] == "BAR"}
    if before.keys() != after.keys():
        raise ValueError("Generated name mapping requires the same physical bar IDs")
    aliases, used_before, used_after, proof = {}, set(), set(), []
    for key, left in before.items():
        right = after[key]
        old, new = decode(left[column]), decode(right[column])
        if not isinstance(old, str) or not isinstance(new, str) or not old or not new:
            raise ValueError(f"Missing generated provenance name: {key}")
        old_key, new_key = (key[0], old), (key[0], new)
        if old_key in used_before or new_key in used_after:
            raise ValueError(f"Ambiguous generated provenance name: {key}")
        used_before.add(old_key)
        used_after.add(new_key)
        aliases[new_key] = old
        if old != new:
            proof.append({"fixture": key[0], "bar_id": key[2], "before": old, "after": new})
    mapped = [row[:] for row in actual]
    for row in mapped:
        if row[1] == "BAR":
            row[column] = before[tuple(row[:3])][column]
        elif row[1] == "LAYOUT":
            key = (row[0], row[2])
            if key not in aliases:
                raise ValueError(f"Layout has no corresponding physical bar: {key}")
            row[2] = aliases[key]
    return mapped, proof


def mutation_tests(records):
    checks = []
    for kind, field in (("CONCRETE", "x"), ("BAR", "diameter"), ("CONTOUR", "role")):
        changed = [row[:] for row in records]
        index = next(i for i, row in enumerate(changed) if row[1] == kind)
        column = 3 + FIELDS[kind].index(field)
        if field == "role":
            changed[index][column] = "s:OPENING" if decode(changed[index][column]) != "OPENING" else "s:OUTER"
        else:
            changed[index][column] = "d:" + struct.pack(">d", decode(changed[index][column]) + 1.0).hex().upper()
        failures = compare(records, changed)
        if len(failures) != 1 or failures[0].get("field") != field:
            raise AssertionError(f"Mutation not detected exactly: {kind}.{field}; {failures}")
        checks.append({"mutation": f"{kind}.{field}", "status": "PASS"})
    if decode("d:3FD5555555555555") != 1.0 / 3.0:
        raise AssertionError("Double decoder is not exact")
    if decode("s:a%25b%09c%0D%0Ad") != "a%b\tc\r\nd":
        raise AssertionError("String decoder lost escaping")
    return checks


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("baseline", type=Path)
    parser.add_argument("actual", type=Path)
    parser.add_argument("--exact", action="store_true")
    parser.add_argument("--skip-input", action="store_true")
    parser.add_argument("--fixture", action="append", default=[], help="Compare only these explicitly selected fixtures")
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--map-generated-rebar-names", action="store_true", help="Record a bijection of bar source/layout names; compare every geometric and physical field unchanged")
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    expected, actual = read_snapshot(args.baseline), read_snapshot(args.actual)
    if args.fixture:
        requested = set(args.fixture)
        for label, records in (("baseline", expected), ("actual", actual)):
            missing = requested - {row[0] for row in records}
            if missing:
                raise ValueError(f"Missing requested fixtures in {label}: {sorted(missing)}")
        expected = [row for row in expected if row[0] in requested]
        actual = [row for row in actual if row[0] in requested]
    name_mapping = []
    if args.map_generated_rebar_names:
        actual, name_mapping = map_generated_rebar_names(expected, actual)
    errors = compare(expected, actual, args.exact, args.skip_input)
    payload = {
        "status": "FAIL" if errors else "PASS", "baseline": str(args.baseline), "actual": str(args.actual),
        "exact": args.exact, "skip_input": args.skip_input, "records": len(actual),
        "requested_fixtures": args.fixture,
        "generated_rebar_name_mapping": name_mapping,
        "fixtures": sorted({row[0] for row in actual}), "errors": errors,
        "mutation_tests": mutation_tests(expected) if args.self_test else [],
    }
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.report:
        args.report.write_text(text + "\n", encoding="utf-8")
    print(json.dumps({"status": payload["status"], "records": len(actual), "errors": len(errors), "mutations": len(payload["mutation_tests"])}))
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
