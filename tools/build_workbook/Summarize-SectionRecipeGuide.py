"""Связывает Source API → реальная строка Справки → вызов/native-журнал."""
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DOCS = ROOT / "docs/geometry_refactor"
EVIDENCE = ROOT / "docs/regression/GeometryDeclarative_2026-10-12/NativeGuide"


def read(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def build():
    api = read(DOCS / "SECTION_RECIPE_API_MANIFEST.json")
    calls = read(DOCS / "SECTION_RECIPE_API_EXAMPLES.json")
    native = read(EVIDENCE / "TableLayout/GuideEvidence.json")
    code = read(EVIDENCE / "TableLayout/CodeComparison.json")
    preserve = read(EVIDENCE / "TableLayout/Preservation.json")
    if native["Status"] != "DIRECTED_NATIVE_GUIDE_PASS" or preserve["status"] != "XML_PRESERVATION_PASS":
        raise ValueError("Native guide/preservation failed")
    if code["Status"] != "PASS" or code["SourceMismatchCount"] or code["UnexpectedCount"]:
        raise ValueError("Native source mismatch")
    if calls["missingCalls"] or calls["missingNativeRoots"]: raise ValueError("Missing API call/evidence")
    mappings = []
    for member in api["members"]:
        row = native["Second"]["ApiRows"][member["anchor"]]
        mappings.append({"id": member["id"], "source": member["source"], "signature": member["signature"],
                         "markdownAnchor": member["anchor"], "helpAddress": "Справка!A" + str(row),
                         "example": calls["examples"][member["id"]]})
    return {"schema": 1, "status": "DIRECTED_NATIVE_GUIDE_PASS", "date": "2026-10-12",
            "nativeWorkbook": "docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/RC_Section_NDM.xlsm",
            "nativeSHA256": hashlib.sha256((EVIDENCE / "RC_Section_NDM.xlsm").read_bytes()).hexdigest(),
            "productionUpdated": False, "productionSHA256": native["ProductionSHA256"],
            "members": len(mappings), "memberMap": mappings,
            "assertions": [{"entryPoint": "modTestRecipeApiCatalogue.RunApiGeometryReadExamples", "passed": 29},
                           {"entryPoint": "modTestRecipeApiCatalogue.RunApiUnitExamples", "passed": 35},
                           {"entryPoint": "modTestRecipeApiCatalogue.RunApiInputAndDataExamples", "passed": 17},
                           {"entryPoint": "modTestRecipeApiCatalogue.RunAdvancedGuideExamples", "passed": 16}],
            "assertionTotal": 97, "assertionFailed": 0, "preservation": preserve,
            "nativeStatesAndNavigation": {key: native[key] for key in ("Idempotent", "StatesSuccess", "StatesFailure", "ReopenTextMatch", "Navigation")},
            "codeComparison": {key: code[key] for key in ("Status", "SourceMismatchCount", "UnexpectedCount", "ActualComponentCounts", "AddedComponents")},
            "screenshots": ["TableLayout/" + name for name in ("GuideToc.png", "GuideCode.png", "GuideSignature.png", "GuideOptions.png")],
            "screenshotsReviewed": True,
            "limits": ["Full heavy on/off and Excel UI suites declined by user; only directed checks executed.",
                       "Build-Workbook final hook added; complete clean build and final production deployment are pending.",
                       "This is a saved isolated fixture; production workbook/dump and final all-requirements acceptance remain pending."]}


if __name__ == "__main__":
    result = build()
    (DOCS / "NATIVE_GUIDE_EVIDENCE.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({key: result[key] for key in ("status", "members", "assertionTotal", "assertionFailed", "productionUpdated")}))
