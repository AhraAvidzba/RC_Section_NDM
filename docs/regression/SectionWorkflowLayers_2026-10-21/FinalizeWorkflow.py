"""Сводит направленную приёмку независимых команд; полный прогон не заявляет."""
from hashlib import sha256
from pathlib import Path
import json
import re

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]


def read(name):
    return json.loads((HERE / name).read_text(encoding="utf-8-sig"))


def log(name):
    text = (HERE / name).read_text(encoding="utf-8-sig")
    assert "WATCHDOG_COMPLETED: exit=0; source unchanged=True" in text, name
    assert not re.search(r"(?m)^FAIL:|RUNTIME ERROR|failed=[1-9]", text), name
    return text


def main():
    package = read("Package21.json")
    accepted = HERE / "Accepted21.xlsm"
    assert package["Passed"] and package["AllInputAndResultCellsPreserved"]
    assert sha256(accepted.read_bytes()).hexdigest().upper() == package["AcceptedSHA256"]
    assert read("Accepted21Equality.json")["Passed"]
    assert all(item["Passed"] for item in read("Structure21.json"))
    no, yes = log("Final_No21.txt"), log("Final_Yes21.txt")
    assert "SOURCE_SHA256: " + package["AcceptedSHA256"] in no
    assert "SOURCE_SHA256: " + package["AcceptedSHA256"] in yes
    for total in ["TOTAL_WORKFLOW_SNAPSHOT: passed=151; failed=0", "TOTAL_WORKFLOW_UI: passed=17; failed=0",
                  "TOTAL_WORKFLOW_CALCULATION: passed=89; failed=0", "TOTAL_AUDIT03_AUTO_PLOT: passed=86; failed=0"]:
        assert total in no, total
    assert "TOTAL_WORKFLOW_CALCULATION: passed=89; failed=0" in yes
    assert re.search(r"RESULTS_SAVE_REOPEN:.*equal=True", yes)
    for name in ["Equivalence_No.json", "Equivalence_Yes.json", "Runtime14/Formatting.json", "NativeCAD11/Stage.json"]:
        assert read(name)["Passed"], name
    assert read("FinalClassComments.json")["status"] == "PASS"
    source = (ROOT / "src/Excel/modWorkbookCalculation.bas").read_text(encoding="utf-8-sig")
    body = source.split("Public Function RunSectionCalculationForWorkbook", 1)[1].split("End Function", 1)[0]
    forbidden = ["BuildWorkbookSectionModel", "FormSectionForWorkbook", "New CSectionBuilder", ".ImportFrom", "CAutoCADSectionModelImporter"]
    assert not any(token in body for token in forbidden)
    assert "ReadFormedSectionFromResults(workbook)" in body
    report = {
        "Passed": True, "Scope": "Directed section workflow acceptance; not full ON/OFF acceptance",
        "AcceptedSHA256": package["AcceptedSHA256"],
        "NoAutomaticGenerationOrImportDuringCalculation": True,
        "SnapshotTests": 151, "CommandTests": 17, "CalculationTestsEachMode": 89, "AutoPlotTests": 86,
        "SaveReopenEqual": True, "SourceAndWorkbookEqual": True,
        "NumericalEquivalence": {mode: read("Equivalence_" + mode + ".json") for mode in ["No", "Yes"]},
        "NativeAutoCAD": read("NativeCAD11/Stage.json"),
        "Formatting": read("Runtime14/Formatting.json"),
        "ReadPerformanceEvidence": "Directed_Off_08.txt; 1000 bars median 0.0625 s; 5000 bars median 0.296875 s",
        "PayloadTablesPerRead": 5, "BoundsScan": "Blocks of up to 1024 values, not individual cells",
        "FullOnOffRunPerformed": False,
        "CodeChangesAfterEquivalenceAndNativeCAD": "Execution report labels and diagnostic comment only; test inputs explicitly reset",
    }
    (HERE / "Acceptance21.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"Passed": True, "Scope": report["Scope"], "AcceptedSHA256": report["AcceptedSHA256"]}))


if __name__ == "__main__":
    main()
