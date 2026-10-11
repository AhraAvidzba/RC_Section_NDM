"""Independent analytic oracle: HollowRectangle paths and normal projections.

Geometry is derived only from baseline INPUT. Concrete/contour records and
all non-coordinate fields retain the original strict snapshot comparison.
"""

import argparse
import copy
import importlib.util
import json
import math
import sys
from pathlib import Path


_spec = importlib.util.spec_from_file_location(
    "rounded_oracle", Path(__file__).with_name("Verify-RoundedRecipeOracle.py")
)
geometry = importlib.util.module_from_spec(_spec)
sys.modules[_spec.name] = geometry
_spec.loader.exec_module(geometry)
snapshots = geometry.snapshots


def cross(a, b):
    return a[0] * b[1] - a[1] * b[0]


def ray_hit(segments, source, direction):
    """Positive ray versus exact finite lines/arcs, nearest admissible hit."""
    hits = []
    for segment in segments:
        if segment.center is None:
            vector = geometry.subtract(segment.end, segment.start)
            denominator = cross(direction, vector)
            if abs(denominator) < 1e-12:
                continue
            delta = geometry.subtract(segment.start, source)
            ray = cross(delta, vector) / denominator
            fraction = cross(delta, direction) / denominator
            if ray >= -1e-7 and -1e-7 <= fraction <= 1 + 1e-7:
                hits.append((max(0, ray), *segment.at(max(0, min(1, fraction)))))
            continue
        delta = geometry.subtract(source, segment.center)
        radius = geometry.norm(geometry.subtract(segment.start, segment.center))
        linear = 2 * sum(a * b for a, b in zip(delta, direction))
        constant = sum(value * value for value in delta) - radius * radius
        discriminant = linear * linear - 4 * constant
        if discriminant < -1e-7:
            continue
        start = math.atan2(segment.start[1] - segment.center[1], segment.start[0] - segment.center[0])
        for ray in ((-linear - math.sqrt(max(0, discriminant))) / 2,
                    (-linear + math.sqrt(max(0, discriminant))) / 2):
            if ray < -1e-7:
                continue
            point = geometry.add(source, geometry.scale(direction, max(0, ray)))
            angle = math.atan2(point[1] - segment.center[1], point[0] - segment.center[0])
            travel = ((angle - start) if segment.sweep > 0 else (start - angle)) % (2 * math.pi)
            if travel > 2 * math.pi - 1e-10:
                travel = 0
            if travel <= abs(segment.sweep) + 1e-10:
                hits.append((max(0, ray), *segment.at(min(1, travel / abs(segment.sweep)))))
    return min(hits, key=lambda hit: hit[0])[1:] if hits else None


def path_for(width, height, radius, center, distance, face, covers):
    x, y = center
    if radius <= 1e-6 and covers is not None:
        top, bottom, left, right = covers
        if face < 2:
            y += (1 if face == 0 else -1) * (height / 2 - distance)
            return [geometry.Segment((x - width / 2 - left, y), (x + width / 2 + right, y))]
        x += (-1 if face == 2 else 1) * (width / 2 - distance)
        return [geometry.Segment((x, y - height / 2 - bottom), (x, y + height / 2 + top))]
    vertices = [(x - width / 2, y + height / 2), (x + width / 2, y + height / 2),
                (x + width / 2, y - height / 2), (x - width / 2, y - height / 2)]
    shifted = geometry.offset_vertices(vertices, distance)
    corners, lines = geometry.fillets(shifted, [max(0, radius - distance)] * 4)
    if face == 0:
        return [segment for segment in (corners[0][2], lines[0], corners[1][2]) if segment is not None and segment.length > 1e-6]
    if face == 1:
        return [segment.reverse() for segment in (corners[3][2], lines[2], corners[2][2]) if segment is not None and segment.length > 1e-6]
    return [lines[3] if face == 2 else lines[1].reverse()]


def expected_geometry(records):
    inputs = {row[2]: snapshots.decode(row[3]) for row in records if row[1] == "INPUT"}
    factor = {"mm": 1.0, "cm": 10.0, "m": 1000.0}[inputs["rngUnitSettings.2.2"]]

    def cell(row, column):
        return inputs.get(f"rngHollowRectangleGeometry.{row}.{column}")

    def length(row, column):
        return float(cell(row, column) or 0) * factor

    outer = length(7, 2), length(7, 1), length(7, 3), (0, 0)
    opening = length(7, 5), length(7, 4), length(7, 6), (length(3, 2), length(4, 2))
    rows = [13, 14, 11, 12, 17, 18, 15, 16]
    groups = ["B.Top", "B.Bottom", "H.Left", "H.Right"]
    covers = [length(row, 2) for row in rows[4:]]
    sources = [[], [], [], []]
    bars, anchors, proofs = [], {}, []

    def available(point):
        return not any(abs(point[0] - bar[0][0]) <= 1e-6 and abs(point[1] - bar[0][1]) <= 1e-6 for bar in bars)

    for index, row in enumerate(rows):
        face = index % 4
        group = ("Opening." if index >= 4 else "") + groups[face]
        diameter = length(row, 3)
        count = int(cell(rows[face], 4))
        if not diameter or not count or not length(rows[face], 3):
            continue
        axis_distance = length(row, 2)
        shape = opening if index >= 4 else outer
        segments = path_for(*shape, -axis_distance if index >= 4 else axis_distance, face,
                            covers if index >= 4 else None)
        total = sum(segment.length for segment in segments)
        proofs.append({"group": group, "length": total, "arc_count": sum(s.center is not None for s in segments)})
        normal_sign = -1 if face in (0, 2) else 1
        points = []
        if index < 4:
            for ordinal in range(1, count + 1):
                fraction = ordinal / (count + 1) if face < 2 else (0.5 if count == 1 else (ordinal - 1) / (count - 1))
                point, tangent = geometry.at_distance(segments, total * fraction)
                points.append((ordinal, point, tangent))
        else:
            for source, direction in sources[face]:
                hit = ray_hit(segments, source, direction)
                if hit:
                    points.append((len(points) + 1, *hit))
            normal_sign *= -1
        main = []
        d2 = length(row + 10, 2)
        for ordinal, point, tangent in points:
            if not available(point):
                continue
            main.append(point)
            direction = geometry.scale((-tangent[1], tangent[0]), normal_sign)
            if index < 4:
                sources[face].append((point, direction))
            bars.append((point, diameter, group, 1))
            for column, extra_row in ((2, 2), (5, 3)):
                extra_d = length(row + 10, column)
                if not extra_d or (cell(row + 10, column + 2) == "EverySecondBar" and ordinal % 2 == 0):
                    continue
                location = cell(row + 10, column + 1)
                distance = (diameter + extra_d) / 2
                if extra_row == 3 and d2 and location == cell(row + 10, 3):
                    distance += d2
                extra = geometry.add(point, geometry.scale(direction if location == "Stacked" else tangent, distance))
                if available(extra):
                    bars.append((extra, extra_d, group, extra_row))
        if face < 2 and main:
            extreme = (max if face == 0 else min)(point[1] for point in main)
            selected = [point for point in main if abs(point[1] - extreme) <= 1e-6]
            anchors[group] = (min(point[0] for point in selected), extreme, max(point[0] for point in selected), extreme)
        elif face >= 2:
            anchors[group] = (*segments[0].start, *segments[-1].end)
    return bars, anchors, proofs


def verify(expected, actual):
    mapped, names = snapshots.map_generated_rebar_names(expected, actual)
    actual_properties = {tuple(row[:3]): row for row in mapped if row[1] == "PROPS"}
    adjusted = copy.deepcopy(expected)
    proofs = []
    for fixture in sorted({row[0] for row in expected}):
        rows = [row for row in expected if row[0] == fixture]
        bars, anchors, paths = expected_geometry(rows)
        source_names = {row[2]: snapshots.decode(row[9]) for row in rows if row[1] == "BAR"}
        if len(bars) != len(source_names):
            raise AssertionError(f"{fixture}: oracle count {len(bars)} != baseline {len(source_names)}")
        layouts = {source_names[f"R{i}"]: bar for i, bar in enumerate(bars, 1)}
        for row in adjusted:
            if row[0] != fixture:
                continue
            if row[1] in ("BAR", "LAYOUT"):
                bar = bars[int(row[2][1:]) - 1] if row[1] == "BAR" else layouts[row[2]]
                if snapshots.decode(row[5]) != bar[1]:
                    raise AssertionError(f"{row[:3]}: oracle diameter/order differs")
                if row[1] == "LAYOUT" and snapshots.decode(row[8]) != bar[2]:
                    raise AssertionError(f"{row[:3]}: oracle owner differs")
                row[3:5] = [geometry.token(value) for value in bar[0]]
            elif row[1] == "ANCHOR" and row[2] in anchors:
                row[3:7] = [geometry.token(value) for value in anchors[row[2]]]
            elif row[1] == "ANNOTATION" and row[2].startswith("REBAR_"):
                row[4:8] = [geometry.token(value) for value in anchors[row[2][6:]]]
        equivalent_axes = []
        angle_column = 3 + snapshots.FIELDS["PROPS"].index("angle")
        for row in adjusted:
            if row[0] != fixture or row[1] != "PROPS":
                continue
            candidate = actual_properties.get(tuple(row[:3]))
            if candidate is None:
                continue
            before, after = snapshots.decode(row[angle_column]), snapshots.decode(candidate[angle_column])
            # An undirected principal axis is periodic in pi, not in 2*pi.
            # Only the exact opposite representation is accepted; tensors stay strict.
            if abs(abs(after - before) - math.pi) <= 1e-12:
                equivalent_axes.append({"expected": before, "actual": after,
                                        "sin_difference": math.sin(after - before)})
                row[angle_column] = candidate[angle_column]
        proofs.append({"fixture": fixture, "bars": len(bars), "paths": paths,
                       "equivalent_principal_axes": equivalent_axes})
    return snapshots.compare(adjusted, mapped), proofs, names


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("baseline", type=Path)
    parser.add_argument("actual", type=Path)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    expected, actual = snapshots.read_snapshot(args.baseline), snapshots.read_snapshot(args.actual)
    errors, proofs, names = verify(expected, actual)
    mutations = []
    probes = (("BAR", "x"), ("BAR", "y"), ("CONTOUR", "role"), ("PROPS", "angle"))
    for kind, field in (() if errors else probes):
        changed = copy.deepcopy(actual)
        index = next(i for i, row in enumerate(changed) if row[1] == kind)
        column = 3 + snapshots.FIELDS[kind].index(field)
        changed[index][column] = "s:OPENING" if field == "role" else geometry.token(snapshots.decode(changed[index][column]) + (0.01 if field == "angle" else 1))
        failures, _, _ = verify(expected, changed)
        if len(failures) != 1 or failures[0].get("field") != field:
            raise AssertionError(f"Oracle mutation not detected exactly: {kind}.{field}")
        mutations.append({"mutation": f"{kind}.{field}", "status": "PASS"})
    result = {"status": "FAIL" if errors else "PASS", "records": len(expected), "errors": errors,
              "fixtures": proofs, "generated_name_mapping": names, "mutation_tests": mutations}
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": result["status"], "fixtures": len(proofs), "records": len(expected), "errors": len(errors)}))
    return bool(errors)


if __name__ == "__main__":
    raise SystemExit(main())
