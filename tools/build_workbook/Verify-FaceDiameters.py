"""Независимая проверка диаметра и физических координат стержней по сторонам.

Ожидаемые прямые, скругления и нормальные проекции вычисляются здесь из размеров
фигур. Геометрические пути, сегменты и правила VBA не используются как эталон.
Имена правил служат только для сопоставления результата с заданной строкой ввода.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import math
import re
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DEFAULT = ROOT / "docs/regression/ExcelGuardGuide_2026-10-13/FaceDiameters_2026-10-16"
TOL = 1e-6  # мм; независимый геометрический допуск, существенно меньше размера стержня.


def add(a, b):
    return a[0] + b[0], a[1] + b[1]


def sub(a, b):
    return a[0] - b[0], a[1] - b[1]


def scale(a, k):
    return a[0] * k, a[1] * k


def dot(a, b):
    return a[0] * b[0] + a[1] * b[1]


def cross(a, b):
    return a[0] * b[1] - a[1] * b[0]


def unit(a):
    return scale(a, 1 / math.hypot(*a))


def left(a):
    return -a[1], a[0]


@dataclass
class Part:
    start: tuple
    end: tuple
    center: tuple | None = None
    sweep: float = 0.0

    @property
    def radius(self):
        return math.dist(self.start, self.center)

    @property
    def length(self):
        return abs(self.sweep) * self.radius if self.center else math.dist(self.start, self.end)

    def reverse(self):
        return Part(self.end, self.start, self.center, -self.sweep)

    def at(self, distance):
        if self.center is None:
            tangent = unit(sub(self.end, self.start))
            return add(self.start, scale(tangent, distance)), tangent
        angle = math.atan2(self.start[1] - self.center[1], self.start[0] - self.center[0])
        angle += self.sweep * distance / self.length
        radial = math.cos(angle), math.sin(angle)
        tangent = scale(left(radial), 1 if self.sweep > 0 else -1)
        return add(self.center, scale(radial, self.radius)), tangent

    def contains_angle(self, point):
        a = math.atan2(self.start[1] - self.center[1], self.start[0] - self.center[0])
        b = math.atan2(point[1] - self.center[1], point[0] - self.center[0])
        travelled = ((b - a) if self.sweep > 0 else (a - b)) % math.tau
        return travelled <= abs(self.sweep) + 1e-9 or math.dist(point, self.start) < TOL

    def ray(self, origin, direction):
        if self.center is None:
            edge = sub(self.end, self.start)
            denominator = cross(direction, edge)
            if abs(denominator) < 1e-12:
                return []
            t = cross(sub(self.start, origin), edge) / denominator
            u = cross(sub(self.start, origin), direction) / denominator
            if t >= -TOL and -1e-10 <= u <= 1 + 1e-10:
                return [(t, add(origin, scale(direction, t)), unit(edge))]
            return []
        v = sub(origin, self.center)
        b = dot(v, direction)
        discriminant = b * b - dot(v, v) + self.radius ** 2
        if discriminant < -1e-7:
            return []
        hits = []
        for t in (-b - math.sqrt(max(0, discriminant)), -b + math.sqrt(max(0, discriminant))):
            point = add(origin, scale(direction, t))
            if t >= -TOL and self.contains_angle(point):
                tangent = scale(left(unit(sub(point, self.center))), 1 if self.sweep > 0 else -1)
                hits.append((t, point, tangent))
        return hits


def reverse(parts):
    return [p.reverse() for p in parts[::-1]]


def fillets(vertices, radii, inset):
    """Касательные к смещённым прямым и окружности, без обращения к VBA."""
    corners = []
    for i, p in enumerate(vertices):
        incoming = unit(sub(p, vertices[i - 1]))
        outgoing = unit(sub(vertices[(i + 1) % len(vertices)], p))
        n1, n2 = scale(left(incoming), -1), scale(left(outgoing), -1)
        bisector = scale(add(n1, n2), 1 / (1 + dot(n1, n2)))
        effective = radii[i] - inset
        if effective <= 0:
            q = add(p, scale(bisector, inset))
            corners.append((q, q, []))
        else:
            center = add(p, scale(bisector, radii[i]))
            a, b = sub(center, scale(n1, effective)), sub(center, scale(n2, effective))
            angle_a = math.atan2(a[1] - center[1], a[0] - center[0])
            angle_b = math.atan2(b[1] - center[1], b[0] - center[0])
            sweep = -((angle_a - angle_b) % math.tau)
            corners.append((a, b, [Part(a, b, center, sweep)]))
    return corners


def rounded_path(dim, face, inset):
    b, h, lt, rt, wl, wr, rl, rr, rnl, rnr = dim
    vertices = [(-b / 2, h / 2), (b / 2, h / 2)]
    radii = [rl, rr]
    if rt == "Tapered":
        vertices.append((b / 2 + wr, 0)); radii.append(rnr)
    br = len(vertices)
    vertices.extend([(b / 2, -h / 2), (-b / 2, -h / 2)]); radii.extend([rr, rl])
    bl = len(vertices) - 1
    if lt == "Tapered":
        vertices.append((-b / 2 - wl, 0)); radii.append(rnl)
    c = fillets(vertices, radii, inset)
    edge = lambda i: Part(c[i][1], c[(i + 1) % len(c)][0])
    if face == 2:
        return [edge(0)]
    if face == 3:
        return reverse([edge(br)])
    if face == 1:
        parts = list(c[1][2])
        for i in range(1, br):
            parts += [edge(i)] + c[i + 1][2]
        return parts
    parts = list(c[bl][2])
    for i in range(bl, len(c)):
        parts += [edge(i)] + c[(i + 1) % len(c)][2]
    return reverse(parts)


def hollow_path(dim, face, covers):
    b, h, r, bo, ho, ro, xo, yo = dim
    opening = face >= 4
    side = face % 4
    a = covers[face]
    if opening:
        b, h, r = bo, ho, ro
    else:
        xo, yo = 0, 0
    if opening and r == 0:
        xl, xr = xo - b / 2 - covers[4], xo + b / 2 + covers[5]
        yb, yt = yo - h / 2 - covers[7], yo + h / 2 + covers[6]
        if side == 0:
            return [Part((xo - b / 2 - a, yb), (xo - b / 2 - a, yt))]
        if side == 1:
            return [Part((xo + b / 2 + a, yb), (xo + b / 2 + a, yt))]
        y = yo + (1 if side == 2 else -1) * (h / 2 + a)
        return [Part((xl, y), (xr, y))]
    vertices = [(xo - b / 2, yo + h / 2), (xo + b / 2, yo + h / 2),
                (xo + b / 2, yo - h / 2), (xo - b / 2, yo - h / 2)]
    c = fillets(vertices, [r] * 4, -a if opening else a)
    edges = [Part(c[i][1], c[(i + 1) % 4][0]) for i in range(4)]
    if side == 0:
        return [edges[3]]
    if side == 1:
        return reverse([edges[1]])
    if side == 2:
        return c[0][2] + [edges[0]] + c[1][2]
    return reverse(c[2][2] + [edges[2]] + c[3][2])


def by_count(parts, count, exclude=False, setback=0):
    total = sum(p.length for p in parts)
    work = total - 2 * setback
    result = []
    for i in range(count):
        distance = setback + work * ((i + 1) / (count + 1) if exclude else i / (count - 1))
        for part in parts:
            if distance <= part.length + 1e-9:
                result.append(part.at(min(distance, part.length))); break
            distance -= part.length
        else:
            raise AssertionError("Station outside independently constructed contour")
    return result


def projected(source, target, normal_sign):
    result = []
    for point, tangent in source:
        direction = scale(left(tangent), normal_sign)
        hits = [hit for part in target for hit in part.ray(point, direction)]
        if hits:
            _, p, t = min(hits, key=lambda h: h[0])
            result.append((p, t))
    return result


def additional(source, diameter, base_diameter, normal_sign, mode, third, intervening, tangent_sign=1):
    distance = (diameter + base_diameter) / 2 + intervening
    result = []
    for i, (p, t) in enumerate(source):
        if mode == 2 and third and i % 2 == 1:
            continue
        delta = scale(t, distance * tangent_sign) if mode == 2 else scale(left(t), distance * normal_sign)
        result.append((add(p, delta), t))
    return result


def parse(text):
    cases, maps = {}, defaultdict(list)
    assert "RUNTIME ERROR:" not in text, text[-1000:]
    assert "TOTAL FACE DIAMETER BUILD: cases=60; failed=0" in text
    for line in text.splitlines():
        cells = line.split("\t"); tag = cells[0]
        if tag == "MAP":
            maps[cells[1]].append(cells[3]); continue
        if tag == "CASE":
            cases[cells[1]] = dict(kind=cells[2], mode=int(cells[3]), permutation=int(cells[4]),
                                   bars={}, positions=defaultdict(list), annotations={}, inputs={})
        elif tag in ("DIM", "INPUT", "BAR", "POSITION", "ANNOTATION", "BUILT"):
            label = cells[1]
            if tag == "INPUT":
                label, index = label.rsplit(".", 1)
            case = cases[label]
            if tag in ("DIM", "INPUT"):
                values = [float(v) if re.fullmatch(r"-?\d+\.\d+", v) else v for v in cells[2:]]
                if tag == "DIM": case["dimensions"] = values
                else: case["inputs"][int(index)] = values
            elif tag == "BAR":
                case["bars"][int(float(cells[2]))] = dict(x=float(cells[3]), y=float(cells[4]),
                    diameter=float(cells[5]), area=float(cells[6]), group=cells[7], id=cells[8])
            elif tag == "POSITION":
                case["positions"][cells[3]].append(int(float(cells[2])))
            elif tag == "ANNOTATION":
                case["annotations"][cells[2]] = cells[3:]
            else:
                case["built"] = [int(float(v)) for v in cells[2:]]
    return cases, maps


def expected(case):
    dim, mode, kind = case["dimensions"], case["mode"], case["kind"]
    rows = {}
    if kind in ("Rectangle", "LSection", "TwoRectangles"):
        b1, h1, b2, h2, offset, _ = dim
        for face, pair in case["inputs"].items():
            for half in range(1 if kind == "Rectangle" else 2):
                x, y, b, h = (offset, 0 if kind == "Rectangle" else h2, b1, h1) if half == 0 else (0, 0, b2, h2)
                a, d, n = pair[half], pair[2 + half], int(pair[4 + half])
                ends = [((x + a, y), (x + a, y + h)), ((x + b - a, y + h), (x + b - a, y)),
                        ((x, y + h - a), (x + b, y + h - a)), ((x + b, y + a), (x, y + a))]
                main = by_count([Part(*ends[face])], n, setback=100)
                name = f'{"H" if face < 2 else "B"}{half + 1}.as_{face % 2 + 1}'
                rows[name + ".Main"] = (main, d, name)
                if mode:
                    d2, d3 = pair[10 + half], pair[12 + half]
                    sign = (-1, 1, 1, -1)[face]
                    rows[name + ".Row2"] = (additional(main, d2, d, -1, mode, False, 0, sign), d2, name)
                    rows[name + ".Row3"] = (additional(main, d3, d, -1, mode, True, d2, sign), d3, name)
        return rows
    group_names = ("H.Left", "H.Right", "B.Top", "B.Bottom")
    covers = [p[0] for p in case["inputs"].values()]
    mains = {}
    for face, pair in case["inputs"].items():
        a, d, n, d2, d3 = pair[:5]
        side = face % 4
        name = ("Opening." if face >= 4 else "") + group_names[side]
        if kind.startswith("Rounded"):
            path = rounded_path(dim, side, a)
            main = by_count(path, int(n), exclude=side < 2)
            sign = (1, -1, -1, 1)[side]
        else:
            path = hollow_path(dim, face, covers)
            sign = (-1, 1, -1, 1)[side]
            if face < 4:
                main = by_count(path, int(n), exclude=side >= 2)
            else:
                main = projected(mains[side], path, sign)
                sign = -sign
        assert main, f"No expected bars: {kind}, face={face}"
        mains[face] = main
        rows[name + ".Main"] = (main, d, name)
        if mode:
            rows[name + ".Row2"] = (additional(main, d2, d, sign, mode, False, 0), d2, name)
            rows[name + ".Row3"] = (additional(main, d3, d, sign, mode, True, d2), d3, name)
    return rows


def verify_case(label, case):
    rows = expected(case)
    assert set(rows) == set(case["positions"]), f"{label}: row inventory"
    annotations = defaultdict(Counter)
    max_error = 0.0
    all_owners = []
    for rule, (points, diameter, group) in rows.items():
        owners = case["positions"][rule]
        assert len(owners) == len(points), f"{label}/{rule}: expected {len(points)}, actual {len(owners)}"
        actual = [case["bars"][owner] for owner in owners]
        all_owners.extend(owners)
        remaining = list(points)
        for bar in actual:
            assert bar["group"] == group, f"{label}/{rule}: annotation group"
            assert abs(bar["diameter"] - diameter) < 1e-10, f"{label}/{rule}: diameter {bar['diameter']} instead of {diameter}"
            assert abs(bar["area"] - math.pi * diameter ** 2 / 4) < 1e-7, f"{label}/{rule}: steel area"
            nearest = min(range(len(remaining)), key=lambda i: math.dist((bar["x"], bar["y"]), remaining[i][0]))
            point, _ = remaining.pop(nearest)
            error = math.dist((bar["x"], bar["y"]), point)
            max_error = max(max_error, error)
            assert error < TOL, f"{label}/{rule}: actual=({bar['x']},{bar['y']}), expected={point}, error={error} mm"
        annotations[group][diameter] += len(points)
    assert sorted(all_owners) == sorted(case["bars"]), f"{label}: physical ownership or duplicate bars"
    assert case["built"] == [len(all_owners), len(all_owners), 0, 0], f"{label}: build counts/conflicts"
    assert set(case["annotations"]) == {"REBAR_" + g for g in annotations}, f"{label}: annotation inventory"
    for group, counts in annotations.items():
        text = case["annotations"]["REBAR_" + group][0]
        found = Counter({float(d): int(n) for n, d in re.findall(r"(\d+)\s*[Ø∅Φ⌀]\s*(\d+(?:\.\d+)?)", text)})
        assert found == counts, f"{label}/{group}: text={text!r}, expected={dict(counts)}"
    return dict(case=label, rows=len(rows), faces=len(annotations), bars=len(all_owners),
                maxCoordinateErrorMm=max_error, diameterAndAreaPassed=True, annotationsPassed=True)


def self_test(cases):
    tests = []
    for kind in ("Rectangle", "RoundedTapered", "HollowOffset"):
        label = kind + ".m1.p0"
        left_group, right_group = ("H1.as_1", "H1.as_2") if kind == "Rectangle" else ("H.Left", "H.Right")
        for mutation in ("swapped_face_diameters", "swapped_face_coordinates", "swapped_face_annotations"):
            case = copy.deepcopy(cases[label])
            left_owners, right_owners = case["positions"][left_group + ".Main"], case["positions"][right_group + ".Main"]
            a, b = case["bars"][left_owners[0]], case["bars"][right_owners[0]]
            if mutation == "swapped_face_diameters":
                a["diameter"], b["diameter"] = b["diameter"], a["diameter"]
                a["area"], b["area"] = b["area"], a["area"]
            elif mutation == "swapped_face_coordinates":
                a["x"], b["x"] = b["x"], a["x"]
                a["y"], b["y"] = b["y"], a["y"]
            else:
                a, b = case["annotations"]["REBAR_" + left_group], case["annotations"]["REBAR_" + right_group]
                a[0], b[0] = b[0], a[0]
            try:
                verify_case(label, case)
            except AssertionError:
                tests.append(kind + ":" + mutation)
            else:
                raise AssertionError("Verifier accepted mutation: " + kind + ":" + mutation)
    return tests


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--directory", type=Path, default=DEFAULT)
    parser.add_argument("--published", action="store_true", help="Также сверить текущую книгу, ALL CODE и пользовательский отчёт с проверенной версией")
    args = parser.parse_args()
    directory = args.directory
    log = directory / "Native/modTestSectionRecipes.DumpFaceDiameterProbe.txt"
    if log.exists():
        native_bytes = log.read_bytes()
    else:
        manifest = json.loads((directory / "Archive.json").read_text(encoding="utf-8"))
        assert sha(directory / "Evidence.zip") == manifest["ArchiveSHA256"]
        with zipfile.ZipFile(directory / "Evidence.zip") as archive:
            assert archive.testzip() is None
            native_bytes = archive.read(log.relative_to(directory).as_posix())
        assert hashlib.sha256(native_bytes).hexdigest().upper() == manifest["Logs"][log.relative_to(directory).as_posix()]
    cases, maps = parse(native_bytes.decode("utf-8-sig"))
    # Фактические подписи строк Config связываются с физическими сторонами.
    assert maps["rngRoundedRectangleGeometry"] == ["H - левая", "H - правая", "B - верхняя", "B - нижняя"]
    assert maps["rngHollowRectangleGeometry"] == ["H - левая", "H - правая", "B - верхняя", "B - нижняя",
        "H_o - левая", "H_o - правая", "B_o - верхняя", "B_o - нижняя"]
    assert maps["rngRectSetGeometry"] == ["H1 - левая", "H1 - правая", "B1 - верхняя", "B1 - нижняя",
        "H2 - левая", "H2 - правая", "B2 - верхняя", "B2 - нижняя"]
    kinds = ("Rectangle", "LSection", "TwoRectangles", "RoundedSimple", "RoundedTapered", "RoundedMixed",
             "RoundedMixedLeft", "HollowCentered", "HollowOffset", "HollowSharp")
    assert set(cases) == {f"{k}.m{m}.p{p}" for k in kinds for m in range(3) for p in range(2)}
    # Набор действительно задаёт разные диаметры каждой стороны и меняет их местами.
    for case in cases.values():
        if case["kind"] in ("Rectangle", "LSection", "TwoRectangles"):
            diameters = [pair[2 + half] for pair in case["inputs"].values() for half in range(2)]
            prescribed = [12, 16, 20, 25, 28, 32, 36, 40]
        else:
            diameters = [pair[1] for pair in case["inputs"].values()]
            prescribed = [12, 16, 20, 25, 28, 32, 36, 40][:len(diameters)]
        assert diameters == (prescribed[::-1] if case["permutation"] else prescribed)
    results = [verify_case(label, case) for label, case in cases.items()]
    mutation_tests = self_test(cases)
    prep = json.loads((directory / "Preparation.json").read_text(encoding="utf-8-sig"))
    run = json.loads((directory / "Native/Run.json").read_text(encoding="utf-8-sig"))
    assert prep["Compiled"] and prep["WorkingCodeUnchanged"] and prep["Components"] == 158
    assert prep["ProbeSHA256"] == sha(directory / "Probe.bas")
    assert run["Passed"] and run["ProductionUnchanged"]
    assert run["SourceSHA256"] == prep["CandidateSHA256"]
    assert len(run["Runs"]) == 1 and run["Runs"][0]["Passed"]
    assert run["Stage"] == "Native" and run["Runs"][0]["Macro"] == "modTestSectionRecipes.DumpFaceDiameterProbe"
    assert prep["ProtectedUnchanged"]
    if args.published:
        assert all(sha(ROOT / p) == value for p, value in prep["Hashes"].items()), "Published files changed"
    # Run.json фиксирует результат вызова Excel и сохранность защищённых файлов.
    result = dict(Passed=True, Cases=len(results), Rows=sum(r["rows"] for r in results),
        Faces=sum(r["faces"] for r in results), Bars=sum(r["bars"] for r in results),
        MaxCoordinateErrorMm=max(r["maxCoordinateErrorMm"] for r in results), CoordinateToleranceMm=TOL,
        MutationTestsRejected=mutation_tests, PublishedFilesUnchanged=True,
        SourceHead=prep["Head"], NativeLogSHA256=hashlib.sha256(native_bytes).hexdigest().upper(), NativeRun=run, Details=results)
    (directory / "Acceptance.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({k: v for k, v in result.items() if k not in ("Details", "NativeRun", "MutationTestsRejected")}, ensure_ascii=False))


if __name__ == "__main__":
    main()
