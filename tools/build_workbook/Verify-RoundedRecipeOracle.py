"""Independent line/arc oracle for the migrated PointedSection fixtures.

Only the proven analytic rebar coordinates and their row extents replace
baseline fields. Every other snapshot field retains the original tolerance.
No VBA paths, contour coordinates or generated bar positions feed the oracle.
"""

import argparse
import copy
import importlib.util
import json
import math
import struct
from dataclasses import dataclass
from pathlib import Path


_spec = importlib.util.spec_from_file_location(
    "snapshots", Path(__file__).with_name("Compare-GeometrySnapshots.py")
)
snapshots = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(snapshots)


def add(a, b):
    return a[0] + b[0], a[1] + b[1]


def subtract(a, b):
    return a[0] - b[0], a[1] - b[1]


def scale(a, factor):
    return a[0] * factor, a[1] * factor


def norm(a):
    return math.hypot(*a)


def unit(a):
    return scale(a, 1 / norm(a))


@dataclass
class Segment:
    start: tuple
    end: tuple
    center: tuple = None
    sweep: float = 0.0

    @property
    def length(self):
        if self.center is None:
            return norm(subtract(self.end, self.start))
        return norm(subtract(self.start, self.center)) * abs(self.sweep)

    def reverse(self):
        return Segment(self.end, self.start, self.center, -self.sweep)

    def at(self, fraction):
        if self.center is None:
            tangent = unit(subtract(self.end, self.start))
            return add(self.start, scale(subtract(self.end, self.start), fraction)), tangent
        radius = norm(subtract(self.start, self.center))
        angle = math.atan2(self.start[1] - self.center[1], self.start[0] - self.center[0])
        angle += fraction * self.sweep
        point = add(self.center, (radius * math.cos(angle), radius * math.sin(angle)))
        sign = math.copysign(1, self.sweep)
        return point, (-sign * math.sin(angle), sign * math.cos(angle))


def offset_vertices(vertices, distance):
    """Intersect adjacent offset supporting lines for the clockwise polygon."""
    result = []
    for i, vertex in enumerate(vertices):
        incoming = unit(subtract(vertex, vertices[i - 1]))
        outgoing = unit(subtract(vertices[(i + 1) % len(vertices)], vertex))
        a = add(vertex, scale((incoming[1], -incoming[0]), distance))
        b = add(vertex, scale((outgoing[1], -outgoing[0]), distance))
        cross = incoming[0] * outgoing[1] - incoming[1] * outgoing[0]
        delta = subtract(b, a)
        along = (delta[0] * outgoing[1] - delta[1] * outgoing[0]) / cross
        result.append(add(a, scale(incoming, along)))
    return result


def fillets(vertices, radii):
    corners = []
    for i, vertex in enumerate(vertices):
        radius = radii[i]
        if radius <= 1e-6:
            corners.append((vertex, vertex, None))
            continue
        before = unit(subtract(vertices[i - 1], vertex))
        after = unit(subtract(vertices[(i + 1) % len(vertices)], vertex))
        alpha = math.acos(max(-1, min(1, before[0] * after[0] + before[1] * after[1])))
        length = radius / math.tan(alpha / 2)
        start, end = add(vertex, scale(before, length)), add(vertex, scale(after, length))
        center = add(vertex, scale(unit(add(before, after)), radius / math.sin(alpha / 2)))
        a = math.atan2(start[1] - center[1], start[0] - center[0])
        b = math.atan2(end[1] - center[1], end[0] - center[0])
        sweep = -((a - b) % (2 * math.pi))
        corners.append((start, end, Segment(start, end, center, sweep)))
    lines = [Segment(corner[1], corners[(i + 1) % len(corners)][0]) for i, corner in enumerate(corners)]
    return corners, lines


def at_distance(segments, distance):
    for segment in segments:
        if distance <= segment.length + 1e-6:
            return segment.at(max(0, min(1, distance / segment.length)))
        distance -= segment.length
    raise AssertionError("Oracle station outside the path")


def path_for(vertices, radii, names, group, axis_distance):
    shifted = offset_vertices(vertices, axis_distance)
    reduced = [max(0, radius - axis_distance) for radius in radii]
    corners, lines = fillets(shifted, reduced)
    if group == "B.Top":
        return [lines[0]]
    if group == "B.Bottom":
        return [lines[names.index("BR")].reverse()]
    result = []
    if group == "H.Right":
        index = names.index("TR")
        last = names.index("BR")
        while True:
            if corners[index][2] is not None:
                result.append(corners[index][2])
            if index == last:
                return result
            result.append(lines[index])
            index = (index + 1) % len(names)
    index = names.index("TL")
    last = names.index("BL")
    while True:
        if corners[index][2] is not None:
            result.append(corners[index][2].reverse())
        if index == last:
            return result
        index = (index - 1) % len(names)
        result.append(lines[index].reverse())


def expected_geometry(records):
    inputs = {row[2]: snapshots.decode(row[3]) for row in records if row[1] == "INPUT"}
    factor = {"mm": 1.0, "cm": 10.0, "m": 1000.0}[inputs["rngUnitSettings.2.2"]]

    def cell(row, column):
        return inputs.get(f"rngPointedSectionGeometry.{row}.{column}")

    def length(row, column):
        return float(cell(row, column) or 0) * factor

    width, height = length(3, 2), length(3, 3)
    types = [cell(7, 2), cell(7, 3)]
    vertices = [(-width / 2, height / 2), (width / 2, height / 2)]
    radii = [length(9, 2), length(9, 3)]
    names = ["TL", "TR"]
    if types[1] == "Tapered":
        vertices.append((width / 2 + length(8, 3), 0.0))
        radii.append(length(10, 3))
        names.append("RightNose")
    vertices.extend([(width / 2, -height / 2), (-width / 2, -height / 2)])
    radii.extend([length(9, 3), length(9, 2)])
    names.extend(["BR", "BL"])
    if types[0] == "Tapered":
        vertices.append((-width / 2 - length(8, 2), 0.0))
        radii.append(length(10, 2))
        names.append("LeftNose")

    bars, anchors, path_proofs = [], {}, []
    for group, row, normal_sign in (("B.Top", 15, -1), ("B.Bottom", 16, 1), ("H.Left", 13, 1), ("H.Right", 14, -1)):
        count, diameter = int(cell(row, 4)), length(row, 3)
        if not count or not diameter:
            continue
        axis_distance = length(row, 2)
        segments = path_for(vertices, radii, names, group, axis_distance)
        total = sum(segment.length for segment in segments)
        path_proofs.append({"group": group, "length": total, "arc_count": sum(s.center is not None for s in segments)})
        main = []
        for ordinal in range(1, count + 1):
            fraction = ordinal / (count + 1) if group.startswith("H.") else (0.5 if count == 1 else (ordinal - 1) / (count - 1))
            point, tangent = at_distance(segments, total * fraction)
            main.append(point)
            candidates = [(point, diameter, 1)]
            d2 = length(row + 6, 2)
            for column, extra_row in ((2, 2), (5, 3)):
                extra_d = length(row + 6, column)
                if not extra_d:
                    continue
                binding = cell(row + 6, column + 2)
                if binding == "EverySecondBar" and ordinal % 2 == 0:
                    continue
                location = cell(row + 6, column + 1)
                distance = diameter / 2 + extra_d / 2
                if extra_row == 3 and d2 and location == cell(row + 6, 3):
                    distance += d2
                direction = scale((-tangent[1], tangent[0]), normal_sign) if location == "Stacked" else tangent
                candidates.append((add(point, scale(direction, distance)), extra_d, extra_row))
            for position, d, extra_row in candidates:
                if not any(abs(position[0] - bar[0][0]) <= 1e-6 and abs(position[1] - bar[0][1]) <= 1e-6 for bar in bars):
                    bars.append((position, d, group, extra_row))
        if group.startswith("H."):
            x = min(point[0] for point in main) if group == "H.Left" else max(point[0] for point in main)
            anchors[group] = (x, max(point[1] for point in main), x, min(point[1] for point in main))
        else:
            anchors[group] = (*segments[0].start, *segments[-1].end)
    return bars, anchors, path_proofs


def token(value):
    return "d:" + struct.pack(">d", value).hex().upper()


def verify(expected, actual):
    mapped, names = snapshots.map_generated_rebar_names(expected, actual)
    adjusted = copy.deepcopy(expected)
    fixtures = sorted({row[0] for row in expected})
    proofs = []
    for fixture in fixtures:
        rows = [row for row in expected if row[0] == fixture]
        bars, anchors, paths = expected_geometry(rows)
        source_names = {row[2]: snapshots.decode(row[3 + snapshots.FIELDS["BAR"].index("source")]) for row in rows if row[1] == "BAR"}
        layouts = {source_names[f"R{i}"]: bar for i, bar in enumerate(bars, 1)}
        expected_count = sum(row[1] == "BAR" for row in rows)
        if len(bars) != expected_count:
            raise AssertionError(f"{fixture}: oracle count {len(bars)} != baseline {expected_count}")
        for row in adjusted:
            if row[0] != fixture:
                continue
            if row[1] in ("BAR", "LAYOUT"):
                bar = bars[int(row[2][1:]) - 1] if row[1] == "BAR" else layouts[row[2]]
                if snapshots.decode(row[5]) != bar[1]:
                    raise AssertionError(f"{row[:3]}: oracle diameter/order differs")
                row[3:5] = [token(bar[0][0]), token(bar[0][1])]
            elif row[1] == "ANCHOR" and row[2] in anchors:
                row[3:7] = [token(value) for value in anchors[row[2]]]
            elif row[1] == "ANNOTATION" and row[2].startswith("REBAR_"):
                row[4:8] = [token(value) for value in anchors[row[2][6:]]]
        proofs.append({"fixture": fixture, "bars": len(bars), "paths": paths})
    return snapshots.compare(adjusted, mapped), proofs, names


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("baseline", type=Path)
    parser.add_argument("actual", type=Path)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    expected = snapshots.read_snapshot(args.baseline)
    actual = snapshots.read_snapshot(args.actual)
    errors, proofs, names = verify(expected, actual)
    mutations = []
    for kind, field in (("BAR", "x"), ("BAR", "y"), ("CONTOUR", "role")):
        changed = copy.deepcopy(actual)
        index = next(i for i, row in enumerate(changed) if row[1] == kind and row[0].endswith("case1"))
        column = 3 + snapshots.FIELDS[kind].index(field)
        changed[index][column] = "s:OPENING" if field == "role" else token(snapshots.decode(changed[index][column]) + 1)
        failures, _, _ = verify(expected, changed)
        if len(failures) != 1 or failures[0].get("field") != field:
            raise AssertionError(f"Oracle mutation was not detected exactly: {kind}.{field}")
        mutations.append({"mutation": f"{kind}.{field}", "status": "PASS"})
    result = {"status": "FAIL" if errors else "PASS", "records": len(expected), "errors": errors, "fixtures": proofs, "generated_name_mapping": names, "mutation_tests": mutations}
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": result["status"], "fixtures": len(proofs), "records": len(expected), "errors": len(errors)}))
    return bool(errors)


if __name__ == "__main__":
    raise SystemExit(main())
