"""Independent necessary force bounds. An inconclusive bound is not a proof."""

import argparse
import copy
import csv
import hashlib
import json
from pathlib import Path

import numpy as np


class ForceBounds:
    def __init__(self, model_path):
        self.fibers = []
        self.materials = {}
        for line in model_path.read_text(encoding="utf-8-sig").splitlines():
            fields = dict(part.split("=", 1) for part in line.split("|")[1:])
            if line.startswith("FIBER|"):
                self.fibers.append({**fields, **{key: float(fields[key]) for key in ("x", "y", "area")}})
            elif line.startswith("MATERIAL|"):
                key = fields["mode"], fields["profile"], fields["kind"]
                self.materials.setdefault(key, []).append((float(fields["strain"]), float(fields["stress"])))
        self.length = max(max(abs(f["x"]), abs(f["y"])) for f in self.fibers)
        self.vectors = np.array([[f["area"], f["area"] * f["y"] / self.length,
                                 f["area"] * f["x"] / self.length] for f in self.fibers])
        directions = list(np.eye(3))
        # Facet normals of the zonotope of independently bounded fiber forces.
        for i in range(len(self.vectors)):
            for j in range(i):
                cross = np.cross(self.vectors[i], self.vectors[j])
                size = np.linalg.norm(cross)
                if size > 1e-12 * np.linalg.norm(self.vectors[i]) * np.linalg.norm(self.vectors[j]):
                    directions.append(cross / size)
        directions = np.array(directions)
        self.directions = np.vstack((directions, -directions))
        self.projections = self.directions @ self.vectors.T
        self.cache = {}

    def support(self, mode, profile):
        key = mode, profile
        if key not in self.cache:
            concrete = np.array(self.materials[(*key, "C")])
            steel = np.array(self.materials[(*key, "S")])
            knots = np.unique(np.concatenate((concrete[:, 0], steel[:, 0])))
            # Steel replaces concrete at the SAME strain, not independent strains.
            net = np.interp(knots, steel[:, 0], steel[:, 1]) - np.interp(knots, concrete[:, 0], concrete[:, 1])
            low = np.array([min(concrete[:, 1]) if f["kind"] == "C" else min(net) for f in self.fibers])
            high = np.array([max(concrete[:, 1]) if f["kind"] == "C" else max(net) for f in self.fibers])
            self.cache[key] = np.maximum(self.projections * low, self.projections * high).sum(axis=1)
        return self.cache[key]

    def check(self, mode, profile, n, mx, my):
        target = np.array([n, mx / self.length, my / self.length])
        support = self.support(mode, profile)
        violation = self.directions @ target - support
        index = int(np.argmax(violation))
        tolerance = max(1e-5, 1e-10 * max(np.max(np.abs(support)), np.linalg.norm(target)))
        return {"ImpossibleByBound": bool(violation[index] > tolerance),
                "ViolationN": float(violation[index]), "ToleranceN": float(tolerance),
                "Direction": self.directions[index].tolist(),
                "ProjectedTargetN": float(self.directions[index] @ target), "SupportN": float(support[index])}

    def check_path(self, mode, profile, origin, base):
        origin = np.array([origin[0], origin[1] / self.length, origin[2] / self.length])
        base = np.array([base[0], base[1] / self.length, base[2] / self.length])
        support = self.support(mode, profile)
        offset = self.directions @ origin
        slope = self.directions @ base
        tolerance = max(1e-5, 1e-10 * max(np.max(np.abs(support)), np.linalg.norm(origin)))
        lower, upper = 0.0, float("inf")
        lower_constraint, upper_constraint = None, None
        for index, (start, change, bound) in enumerate(zip(offset, slope, support)):
            constraint = {"Direction": self.directions[index].tolist(), "OriginN": float(start),
                          "BaseN": float(change), "SupportN": float(bound), "ToleranceN": float(tolerance)}
            rhs = bound + tolerance - start
            if change == 0.0:
                if rhs < 0.0:
                    return {"ImpossibleByBound": True, "ConstantConstraint": constraint,
                            "Meaning": "No nonnegative lambda satisfies this necessary force bound."}
            elif change > 0.0:
                limit = float(rhs / change)
                if limit < upper:
                    upper, upper_constraint = limit, constraint
            else:
                limit = float(rhs / change)
                if limit > lower:
                    lower, lower_constraint = limit, constraint
        scale = max(1.0, abs(lower), abs(upper)) if np.isfinite(upper) else float("inf")
        impossible = upper < lower and lower - upper > 1e-12 * scale
        return {"ImpossibleByBound": bool(impossible), "LowerLambda": lower,
                "UpperLambda": upper if np.isfinite(upper) else None,
                "LowerConstraint": lower_constraint, "UpperConstraint": upper_constraint,
                "Meaning": "Necessary interval only; a nonempty interval does not prove equilibrium."}

    def check_physical_formation_path(self, profile, origin, base):
        points = sorted({(f["x"], f["y"]) for f in self.fibers if f["kind"] == "C"})

        def cross(a, b, c):
            return (b[0]-a[0])*(c[1]-a[1]) - (b[1]-a[1])*(c[0]-a[0])

        lower, upper = [], []
        for point in points:
            while len(lower) >= 2 and cross(lower[-2], lower[-1], point) <= 0:
                lower.pop()
            lower.append(point)
        for point in reversed(points):
            while len(upper) >= 2 and cross(upper[-2], upper[-1], point) <= 0:
                upper.pop()
            upper.append(point)
        hull = lower[:-1] + upper[:-1]
        bars = [f for f in self.fibers if f["kind"] == "S"]
        if len(hull) < 3 or any(cross(a, hull[(i+1) % len(hull)], (bar["x"], bar["y"])) < 0
                              for bar in bars for i, a in enumerate(hull)):
            return {"ImpossibleByBound": False, "Unavailable": "Steel is outside the concrete-center hull."}
        # An affine strain at an interior bar is a convex combination of
        # concrete strains. In a PHYSICAL formation state, every bar therefore
        # lies within the concrete compression/tension strain limits as well.
        concrete = np.array(self.materials[("No", profile, "C")])
        steel = np.array(self.materials[("No", profile, "S")])
        minimum, maximum = float(concrete[0, 0]), float(concrete[-1, 0])
        knots = np.unique(np.r_[concrete[:, 0], steel[:, 0], minimum, maximum])
        knots = knots[(knots >= minimum) & (knots <= maximum)]
        bounded = copy.copy(self)
        bounded.materials = self.materials.copy()
        bounded.cache = {}
        bounded.materials[("No", "PhysicalFormation", "C")] = list(zip(knots, np.interp(knots, concrete[:, 0], concrete[:, 1])))
        bounded.materials[("No", "PhysicalFormation", "S")] = list(zip(knots, np.interp(knots, steel[:, 0], steel[:, 1])))
        result = bounded.check_path("No", "PhysicalFormation", origin, base)
        result.update(ConcreteStrainRange=[minimum, maximum], SteelInsideConcreteCenterHull=True,
                      SteelCount=len(bars), ConcreteCenterHull=hull,
                      Meaning="Necessary bound for a physical crack-formation plane, not for an arbitrary extended equilibrium.")
        return result


def path_origin(n, mx, my, path):
    component = path.split("*")[-1]
    if component == "N":
        n = 0.0
    elif component == "Mx":
        mx = 0.0
    elif component == "My":
        my = 0.0
    elif component == "Mxy":
        mx = my = 0.0
    elif component in ("NMxy", "Auto"):
        n = mx = my = 0.0
    else:
        raise ValueError("Unknown load path: " + path)
    return n, mx, my


def audit(rows, model):
    records = []
    for row in rows:
        for stage in ("direct", "capacity", "formation", "current", "width", "longitudinal", "stability"):
            if row[stage + "_InternalStatus"] != "rsNumericalFailure":
                continue
            record = {"Mode": row["ExtensionEnabled"], "Stage": stage, "Load": row["Load"],
                      "Path": row["RequestedPath"], "Code": row[stage + "_ResultCode"],
                      "Comment": row[stage + "_ResultComment"], "CombinationID": row["CombinationID"]}
            record.update(Family=row["Family"], Log=row["Log"])
            if record["Code"] == "SEARCH_BOUND_REACHED":
                record["Classification"] = "TechnicalSearchBound"
                records.append(record)
                continue
            if stage not in ("direct", "capacity", "formation", "current"):
                record["Classification"] = "RequiresNonEquilibriumReview"
                records.append(record)
                continue
            n, mx, my = [float(row[key]) for key in ("N_Internal", "Mx_Internal", "My_Internal")]
            profile = "PR2" if stage == "current" else "PR3" if stage == "formation" else "PR1"
            # The matrix uses +10/-7 mm from the concrete centroid.
            full = model.check(record["Mode"], profile, n, mx - n * 7.0, my + n * 10.0)
            if stage in ("direct", "current"):
                record.update(TargetKind="CurrentLoad", N=n, MxCenter=mx-n*7.0, MyCenter=my+n*10.0, **full)
                record["Classification"] = "ImpossibleCurrentEquilibrium" if full["ImpossibleByBound"] else "Inconclusive"
            else:
                path = row["FormationPath"] if stage == "formation" else row["RequestedPath"]
                if not path:
                    record["Classification"] = "InconclusiveMissingPath"
                    records.append(record)
                    continue
                origin_n, origin_mx, origin_my = path_origin(n, mx, my, path)
                origin = model.check(record["Mode"], profile, origin_n,
                                     origin_mx-origin_n*7.0, origin_my+origin_n*10.0)
                record.update(TargetKind="PathOrigin", N=origin_n, MxCenter=origin_mx-origin_n*7.0,
                              MyCenter=origin_my+origin_n*10.0, **origin)
                record["FullLoadBound"] = full
                record["PathBound"] = model.check_path(record["Mode"], profile,
                    (origin_n, origin_mx-origin_n*7.0, origin_my+origin_n*10.0),
                    (n-origin_n, (mx-origin_mx)-(n-origin_n)*7.0, (my-origin_my)+(n-origin_n)*10.0))
                if stage == "formation" and not record["PathBound"]["ImpossibleByBound"]:
                    record["FormationBound"] = model.check_physical_formation_path(profile,
                        (origin_n, origin_mx-origin_n*7.0, origin_my+origin_n*10.0),
                        (n-origin_n, (mx-origin_mx)-(n-origin_n)*7.0, (my-origin_my)+(n-origin_n)*10.0))
                # This proves absence at lambda=0, not absence on the ENTIRE path.
                record["Classification"] = ("ImpossibleEntirePath" if record["PathBound"]["ImpossibleByBound"]
                    else "ImpossibleFormationCriterion" if record.get("FormationBound", {}).get("ImpossibleByBound")
                    else "ImpossiblePathOrigin" if origin["ImpossibleByBound"] else "Inconclusive")
            records.append(record)
    return records


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--csv", required=True, type=Path)
    parser.add_argument("--model", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    model = ForceBounds(args.model)
    with args.csv.open(encoding="utf-8-sig", newline="") as stream:
        records = audit(list(csv.DictReader(stream)), model)
    csv_hash = hashlib.sha256(args.csv.read_bytes()).hexdigest().upper()
    model_hash = hashlib.sha256(args.model.read_bytes()).hexdigest().upper()
    for record in records:
        record.update(CsvSHA256=csv_hash, ModelSHA256=model_hash)
    args.output.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
    counts = {}
    for record in records:
        key = record["Mode"], record["Stage"], record["Classification"]
        counts[key] = counts.get(key, 0) + 1
    for key, count in sorted(counts.items()):
        print("FORCE_AUDIT:", *key, "count=" + str(count))
    print("INCONCLUSIVE_COUNT:", sum("Inconclusive" in r["Classification"] or r["Classification"] == "ImpossiblePathOrigin" or
                                    r["Classification"].startswith("Requires") for r in records))


if __name__ == "__main__":
    main()
