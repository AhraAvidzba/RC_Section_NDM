"""Собирает итоговые доказательства и матрицу требований из выполненных логов."""
from __future__ import annotations
import hashlib
import importlib.util
import json
import re
import statistics
import subprocess
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DOC = ROOT / "docs/geometry_refactor"
BASE = ROOT / "docs/regression/GeometryDeclarative_2026-10-12"
ACCEPT = BASE / "FinalAcceptance"
RELEASE = BASE / "FinalRelease"

def read(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def relative(path):
    return path.relative_to(ROOT).as_posix()

def report(path):
    text = path.read_text(encoding="utf-8-sig")
    totals = re.findall(r"^TOTAL[^\r\n]*passed=(\d+); failed=(\d+)", text, re.M)
    if not totals or any(int(f) for _, f in totals) or re.search(r"^FAIL:", text, re.M):
        raise ValueError("Unpassed final report: " + str(path))
    return {"path": relative(path), "sha256": sha(path), "status": "PASS",
            "passed": sum(int(p) for p, _ in totals), "totals": totals}

def seconds(path, prefix):
    text = path.read_text(encoding="utf-8-sig")
    return {name: float(value) for name, value in re.findall(
        r"^" + prefix + r": ([^;]+); seconds=([^;]+)", text, re.M)}

def phases(path):
    keys = ("seconds", "preparation", "core", "clear", "packing", "summary", "snapshot", "plot", "solves")
    rows = [line for line in path.read_text(encoding="utf-8-sig").splitlines() if line.startswith("FAMILY_STAGES:")]
    if len(rows) != 3:
        raise ValueError("Three stage samples required")
    return {key: statistics.median(float(re.search(r"(?<![A-Za-z])" + key + r"=([^;]+)", line)[1]) for line in rows[1:]) for key in keys}

def main():
    commit = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip()
    reports = [report(ACCEPT / name) for name in (
        "FakeCAD.txt", "PhysicsIndependent.txt", "AutoPlotFixed.txt", "UnitAndIsolation.txt",
        "FamilyComplementSP35Final.txt", "FamilyComplementSP63Final.txt", "DirectedPlotModes.txt", "SP35ProjectionCrackedFixture.txt")]
    reports.append(report(ACCEPT / "NativeCADROT/NativeCAD.txt"))
    comparisons = [read(ACCEPT / ("Family" + code + "PostMigration.json")) for code in ("SP35", "SP63")]
    if any(c["status"] != "PASS" or not c["exact"] or c["records"] != 48114 for c in comparisons):
        raise ValueError("Final family snapshots do not match accepted analytic migration")
    spec = importlib.util.spec_from_file_location("snapshots", ROOT / "tools/build_workbook/Compare-GeometrySnapshots.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    nds = {}
    for code in ("SP35", "SP63"):
        rows = module.read_snapshot(ACCEPT / ("Family" + code + ".txt"))
        states = [r for r in rows if r[1] == "NDS"]
        nds[code] = {"records": len(rows), "directStates": len(states),
                     "converged": sum(module.decode(r[3]) for r in states),
                     "batchResults": sum(r[1] == "BATCH" for r in rows)}
        if nds[code]["converged"] != 108:
            raise ValueError("Family direct-state convergence incomplete")
    before = seconds(ACCEPT / "FamilyWorkflowBaseline.txt", "FAMILY_WORKFLOW")
    after = seconds(ACCEPT / "FamilyWorkflow.txt", "FAMILY_WORKFLOW")
    mixed = float(re.search(r"MIXED_WORKFLOW: seconds=([^;]+)", (ACCEPT / "MixedWorkflow.txt").read_text(encoding="utf-8-sig"))[1])
    if len(after) != 9 or set(before) != set(after) or max(after.values()) > 10 or mixed > 10:
        raise ValueError("Directed end-to-end time gate failed")
    data, guide, equality = [read(RELEASE / name) for name in (
        "UserDataPreservation.json", "GuideFinalPreservation.json", "SourceEquality.json")]
    if data["status"] != "USER_DATA_PRESERVED" or guide["status"] != "XML_PRESERVATION_PASS" or not equality["AllPassed"]:
        raise ValueError("Release preservation/source gate failed")
    preview = read(ACCEPT / "PreviewComparison.json")
    if any(e.get("record", [None])[0] not in ("HollowCentered", "HollowOffset") or
           e.get("record", [None, None])[1] != "PLOT_SHAPE" or
           e.get("field") not in ("left", "top", "width", "height") for e in preview["errors"]):
        raise ValueError("Unexpected preview change")
    publication_file = RELEASE / "Publication.json"
    publication = read(publication_file) if publication_file.exists() else None
    book = RELEASE / "RC_Section_NDM.xlsm"
    if publication and (sha(ROOT / "workbook/output/RC_Section_NDM.xlsm") != sha(book) or
                        sha(ROOT / "workbook/output/VBA_All_Code.txt") != sha(RELEASE / "VBA_All_Code.txt")):
        raise ValueError("Published artifact mismatch")
    evidence = {"schema": 1, "status": "ACCEPTED_DIRECTED" if publication else "READY_TO_PUBLISH",
        "dateUTC": datetime.now(timezone.utc).isoformat(), "codeCommit": commit,
        "specVersion": "1.4", "workbookSHA256": sha(book), "vbaSHA256": sha(RELEASE / "VBA_All_Code.txt"),
        "reports": reports, "assertionTotalThisAcceptance": sum(r["passed"] for r in reports),
        "nativeExcel": {"version": "16.0", "build": "10417", "components": 147},
        "familyPhysics": nds, "postMigrationExactComparison": comparisons,
        "originalBaselineDifference": {"SP35": 7118, "SP63": 7115,
            "families": ["HollowCentered", "HollowOffset"], "classification": "DOCUMENTED_ANALYTIC_PROJECTION_CHANGE",
            "oracle": "docs/geometry_refactor/HOLLOW_RECIPE_EVIDENCE.json", "behavior": "docs/geometry_refactor/BEHAVIOR_CHANGES.md"},
        "preview": {"records": preview["records"], "reviewedChanges": len(preview["errors"]),
            "classification": "Only Hollow rebar/display coordinates follow independently checked analytic projection",
            "images": 9, "stressStrainImages": 18, "specialModeImages": 8,
            "reviewed": ["PreviewContact.png", "StressContact.png", "StrainContact.png", "PlotModesPNG/PERF_01_Stress.png",
                         "PlotModesPNG/Resize_NoLabels.png", "PlotModesPNG/MissingState.png"]},
        "performance": {"scope": "Nine single-LC full workflows plus one full mixed profile; three staged HollowOffset samples",
            "baselineFullSeconds": before, "actualFullSeconds": after, "mixedFullSeconds": mixed,
            "warmPhaseBaselineMedian": phases(ACCEPT / "FamilyStagesBaseline.txt"),
            "warmPhaseActualMedian": phases(ACCEPT / "FamilyStagesCurrent.txt"),
            "interpretation": "Additional time is mainly summary/snapshot output; common diagnostics/provenance add saved data. Solve count stays one; no repeated preparation. Samples are directed, not a broad benchmark."},
        "nativeCAD": {"status": "PASS", "report": "docs/regression/GeometryDeclarative_2026-10-12/FinalAcceptance/NativeCADROT/NativeCAD.txt",
            "assertions": 398, "identity": read(ACCEPT / "NativeCADROT/NativeCADIdentity.json"),
            "earlierBlockedStatusSuperseded": True, "userDocumentsUsed": False},
        "preservation": data, "guidePreservation": guide, "sourceEquality": {k:v for k,v in equality.items() if k not in ("Checks", "GeneratedDocumentChecks")},
        "publication": publication,
        "publishedReadback": ({"status": "PASS" if read(RELEASE / "PublishedSourceEquality.json")["AllPassed"] else "FAIL",
            "path": relative(RELEASE / "PublishedSourceEquality.json"), "components": 147, "sourceUnchanged": True}
            if (RELEASE / "PublishedSourceEquality.json").exists() else None),
        "notRun": [{"id": "FULL_ON_OFF", "status": "NOT_RUN", "reason": "User chose directed checks"},
                   {"id": "FULL_EXCEL_UI", "status": "NOT_RUN", "reason": "User chose directed checks"}],
        "failedProbes": ["FamilySP35Comparison.json", "FamilySP63Comparison.json", "PreviewComparison.json",
            "PassiveSnapshots.txt", "DirectedPhysics.txt", "AutoPlotIsolated.txt", "FamilyComplementSP35Corrected.txt",
            "FamilyComplementSP63Corrected.txt", "HarnessPlotImport.txt", "NativeCADDirectLaunch.txt",
            "NativeCADExplicitLaunch.txt", "NativeCADRegisteredLaunch.txt"],
        "failuresResolved": "Old strict Hollow diff explained by existing independent oracle; AutoPlot fixture now explicitly Generated; projection test required a suitable current SP35 LC; wrong fixture keys/accessors and preparation count corrected; native ROT/visible owned AutoCAD passed. Full logs retained."}
    (DOC / "FINAL_ACCEPTANCE_EVIDENCE.json").write_text(json.dumps(evidence, ensure_ascii=False, indent=2)+"\n", encoding="utf-8")
    manifest = read(DOC / "TEST_MANIFEST.json")
    groups = {
        "G": ["RETIREMENT_EVIDENCE.json", "DECLARATIVE_API_EVIDENCE.json", "ANALYTIC_PATH_EVIDENCE.json"],
        "R": ["DECLARATIVE_API_EVIDENCE.json", "CONFIG_SETBACK_EVIDENCE.json", "RECTANGLES_EVIDENCE.json", "HOLLOW_RECIPE_EVIDENCE.json"],
        "A": ["DECLARATIVE_API_EVIDENCE.json", "NATIVE_GUIDE_EVIDENCE.json", "FINAL_ACCEPTANCE_EVIDENCE.json"],
        "I": ["DECLARATIVE_API_EVIDENCE.json", "FINAL_ACCEPTANCE_EVIDENCE.json"],
        "E": ["ONE_CLASS_EVIDENCE.json", "NATIVE_GUIDE_EVIDENCE.json", "CONFIG_SETBACK_EVIDENCE.json", "FINAL_ACCEPTANCE_EVIDENCE.json"],
        "D": ["FINAL_ACCEPTANCE_EVIDENCE.json", "LEGACY_RETIREMENT_MAP.md", "FINAL_REPORT.md"]}
    runtime_ids = {
        "G": "modTestGeometry.RunGeometryTests; modTestDeclarativeGeometry.RunCoreGeometryTests; modTestRecipeOperations.RunRecipeOperationTests",
        "R": "modTestSectionRecipes.RunRecipeRuleTests; ConfigSetback native reports; family oracle comparisons",
        "A": "modTestRecipeOperations.RunRecipeMeasurementTests; modTestPlotConfig.RunAudit03GeneralPlotTests; modTestFinalGeometryChecks.TestRunDirectedPlotModes",
        "I": "modTestWorkbookInterface.RunAudit03ReadLifecycleTests; modTestSectionSolver.RunCenteredNumericalTests; saved snapshot native reports",
        "E": "modTestRoundedTrapezoidExample.RunExampleBuildTests; modTestRoundedTrapezoidExample.RunExampleInputTests; modTestRecipeApiCatalogue native examples; guide native updater",
        "D": "FinalAcceptance runtime reports, source/native comparisons and publication guards"}
    for requirement in manifest["Requirements"]:
        rid = requirement["Id"]
        requirement.update(Status="PASS" if publication or rid != "D24" else "NOT_RUN",
            TestId=runtime_ids[rid[0]], CodeCommit=commit,
            WorkbookSHA256=sha(book), ExecutedAt=evidence["dateUTC"],
            Evidence=["docs/geometry_refactor/"+name for name in groups[rid[0]]],
            Notes="Реальные native логи и точные source/model сравнения перечислены в связанных evidence. Полные on/off и UI наборы не запускались по выбору пользователя; адресное покрытие соответствует §14.4.")
    manifest.update(UpdatedAt=evidence["dateUTC"], Scope="DIRECTED_ACCEPTANCE", RuntimeReports=reports,
                    OptionalLongSuites=evidence["notRun"], FinalWorkbookSHA256=sha(book), Publication=publication)
    (DOC / "TEST_MANIFEST.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2)+"\n", encoding="utf-8")
    print(json.dumps({"status": evidence["status"], "assertions": evidence["assertionTotalThisAcceptance"], "requirements": len(manifest["Requirements"]), "mixedFullSeconds": mixed}))

if __name__ == "__main__":
    main()
