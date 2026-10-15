"""Проверяет каждый фактический NumFail независимой границей усилий.

Использует общий проверочный ForceBounds из сохранённого аудита. Невозможность
усилий доказывается в расширенном множестве независимых напряжений элементов;
непустое множество не доказывает существования плоского НДС. Статусы VBA
не изменяются, код этой проверки в рабочую книгу не попадает.
"""

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
BOUNDS_SOURCE = ROOT / "docs/regression/PostAudit03/Axial1000_2026-10-09/AuditForceBounds.py"
sys.path.insert(0, str(BOUNDS_SOURCE.parent))
from AuditForceBounds import ForceBounds


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def fields(line):
    return dict(part.split("=", 1) for part in line.split("|")[1:] if "=" in part)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--directory", default="docs/regression/Performance/CoreRegression_2026-10-15")
    args = parser.parse_args()
    directory = ROOT / args.directory
    inputs = json.loads((directory / "InputsBefore.json").read_text(encoding="utf-8-sig"))
    manifest = json.loads((directory / "ForceBounds/ModelManifest.json").read_text(encoding="utf-8-sig"))
    model_path = directory / "ForceBounds/Model.txt"
    if manifest["SourceSHA256"] != inputs["CurrentSHA256"].upper() or digest(model_path) != manifest["DataSHA256"] or not inputs["DataPartsIdentical"]:
        raise ValueError("Force model provenance does not match the comparison.")
    model = ForceBounds(model_path)
    if len(model.fibers) != 65 or len(model.materials) != 12:
        raise ValueError("Incomplete saved model or material roles.")
    model_lines = model_path.read_text(encoding="utf-8-sig").splitlines()
    tolerance = fields(next(line for line in model_lines if line.startswith("ACCEPTANCE_TOLERANCE|")))
    center = fields(next(line for line in model_lines if line.startswith("CENTER|")))
    tn, tmx, tmy = (float(tolerance[key]) for key in ("N", "Mx", "My"))
    # Консервативно учитываем и допуски исходных осей, и перенос к центру.
    centered_tmx = tmx + abs(float(center["y"])) * tn
    centered_tmy = tmy + abs(float(center["x"])) * tn
    loads = {}
    for line in (directory / "After/SavedWorkbook_Saved_Manifest.txt").read_text(encoding="utf-8-sig").splitlines():
        if line.startswith("FOCUSED_INPUT|"):
            load = fields(line)
            loads[load["id"]] = tuple(float(load[key]) for key in ("N", "Mx", "My"))
    progress = json.loads((directory / "Progress.json").read_text(encoding="utf-8-sig"))
    certificates, checked_meta, unsupported = [], [], []
    for report in progress:
        path = ROOT / report["Report"]
        if digest(path) != report["ReportSHA256"]:
            raise ValueError("Changed report: " + str(path))
        lines = path.read_text(encoding="utf-8-sig").splitlines()
        states = {}
        for line in lines:
            if not line.startswith("FOCUSED_STATE|"):
                continue
            state = fields(line)
            if state["status"] != "NumFail" or state["converged"] != "False":
                continue
            if report["Shape"] != "SavedWorkbook" or state["type"] not in ("StrengthState", "CrackedState"):
                unsupported.append({"report": report["ID"], "line": line})
                continue
            prefix = "plateau.batch15." if report["Group"] == "Batch15" else "focused.SavedWorkbook."
            if not state["id"].startswith(prefix):
                raise ValueError("Unexpected state ID: " + state["id"])
            load_id = state["id"][len(prefix):]
            n, mx, my = loads[load_id]
            role = "PR1" if state["type"] == "StrengthState" else "PR2"
            certificate = model.check(report["Mode"], role, n, mx - 7.0 * n, my + 10.0 * n)
            direction = certificate["Direction"]
            projection_tolerance = abs(direction[0]) * tn + (abs(direction[1]) * centered_tmx + abs(direction[2]) * centered_tmy) / model.length
            certificate["AcceptanceProjectionN"] = projection_tolerance
            certificate["ProvenBeyondAcceptanceTolerance"] = certificate["ViolationN"] > projection_tolerance + certificate["ToleranceN"]
            states[(load_id, "direct" if role == "PR1" else "current")] = certificate
            certificates.append({"report": report["ID"], "id": load_id, "mode": report["Mode"],
                                 "state": state["type"], "role": role, "N": n,
                                 "MxCenter": mx - 7.0 * n, "MyCenter": my + 10.0 * n,
                                 "certificate": certificate, "reportSHA256": report["ReportSHA256"]})
        for line in lines:
            if not line.startswith("META:") or "|status=rsNumericalFailure|" not in line:
                continue
            match = re.match(r"META: audit03\.comments\.(.+)\.(direct|current)\|", line)
            if not match:
                unsupported.append({"report": report["ID"], "line": line})
                continue
            key = (match[1], match[2])
            checked_meta.append({"report": report["ID"], "id": key[0], "stage": key[1],
                                 "covered": key in states and states[key]["ProvenBeyondAcceptanceTolerance"]})
    output = {
        "ReviewedStates": len(certificates),
        "ProvenImpossibleByNecessaryForceBound": sum(item["certificate"]["ImpossibleByBound"] for item in certificates),
        "ProvenBeyondAcceptanceTolerance": sum(item["certificate"]["ProvenBeyondAcceptanceTolerance"] for item in certificates),
        "Inconclusive": sum(not item["certificate"]["ProvenBeyondAcceptanceTolerance"] for item in certificates),
        "AllNumericalStagesCovered": bool(checked_meta) and not unsupported and all(item["covered"] for item in checked_meta),
        "BoundsSourceSHA256": digest(BOUNDS_SOURCE), "ModelSHA256": digest(model_path),
        "SourceSHA256": manifest["SourceSHA256"], "ProductionPrecheckAdded": False,
        "Scope": "Only full requested NDS targets; not a statement about every lambda on a path.",
        "Unsupported": unsupported, "Stages": checked_meta, "Certificates": certificates,
    }
    (directory / "NumericalFailureReview.json").write_text(json.dumps(output, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({key: output[key] for key in ("ReviewedStates", "ProvenImpossibleByNecessaryForceBound", "Inconclusive", "AllNumericalStagesCovered")}, ensure_ascii=False))
    if not output["AllNumericalStagesCovered"] or output["Inconclusive"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
