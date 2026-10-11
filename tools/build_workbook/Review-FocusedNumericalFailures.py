"""Test-only necessary force certificates; inconclusive bounds are not failures of physics."""
import argparse
import hashlib
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SERIES = ROOT / "docs/regression/PostAudit03/Axial1000_2026-10-09"
sys.path.insert(0, str(SERIES / "ExternalSciPy/vendor"))
sys.path.insert(0, str(SERIES))
from AuditForceBounds import ForceBounds


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def fields(line):
    return dict(part.split("=", 1) for part in line.split("|")[1:] if "=" in part)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--model", type=Path, required=True)
    parser.add_argument("--load-stage", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    manifest = json.loads(args.model.with_name("ModelManifest.json").read_text(encoding="utf-8-sig"))
    stage = json.loads(args.load_stage.read_text(encoding="utf-8-sig"))
    assert manifest["SourceSHA256"] == stage["SourceSHA256"]
    assert digest(args.model) == manifest["DataSHA256"]
    assert digest(Path(manifest["SourceWorkbook"])) == stage["SourceSHA256"]
    model = ForceBounds(args.model)
    records = []
    for group in stage["Groups"]:
        path = ROOT / group["Report"]
        assert digest(path) == group["ReportSHA256"]
        lines = path.read_text(encoding="utf-8-sig").splitlines()
        targets = {}
        for line in lines:
            if line.startswith("CASE: focused.") and "|N=" in line and "|offsetX=" in line:
                data = fields(line)
                key = line.split("|", 1)[0].removeprefix("CASE: ")
                n, mx, my = (float(data[name]) for name in ("N", "Mx", "My"))
                targets[key] = (n, mx + n * float(data["offsetY"]), my + n * float(data["offsetX"]))
        for line in lines:
            if not line.startswith("FOCUSED_STATE|"):
                continue
            state = fields(line)
            if state["converged"] != "False" or state["status"] != "NumFail":
                continue
            assert group["Shape"] == manifest["Shape"], "Export a separate model for the failed section shape."
            assert state["type"] in ("StrengthState", "CrackedState"), "A failed limit path needs its own whole-path proof."
            profile = "PR1" if state["type"] == "StrengthState" else "PR2"
            certificate = model.check(group["Mode"], profile, *targets[state["id"]])
            records.append({"Case": state["id"], "Mode": group["Mode"], "State": state["type"],
                            "MaterialRole": profile, "Certificate": certificate, "ReportSHA256": digest(path)})
    output = {"SourceSHA256": stage["SourceSHA256"], "ModelSHA256": digest(args.model),
              "ReviewedStates": len(records), "ProvenImpossible": sum(r["Certificate"]["ImpossibleByBound"] for r in records),
              "Inconclusive": sum(not r["Certificate"]["ImpossibleByBound"] for r in records),
              "Records": records, "ProductionPrecheckAdded": False,
              "Scope": "Full requested NDS targets, not partial load steps; not a proof about all lambda paths."}
    args.output.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({key: output[key] for key in ("ReviewedStates", "ProvenImpossible", "Inconclusive")}))


if __name__ == "__main__":
    main()
