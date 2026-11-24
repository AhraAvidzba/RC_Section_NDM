"""Проверяет адресную приёмку и переносит только VBA, сохраняя оформление и данные.

Книга собирается из исходного пакета и скомпилированного проекта VBA. Скрипт
не заменяет рабочий файл; подготовленный Accepted.xlsm проверяется до установки.
"""
import argparse
import hashlib
import json
from pathlib import Path
import re
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
EXPECTED_ORIGINAL = "2D413237AAF72D2D134290268F828DBC99140C2850303B558D4B1F68C96AAE06"
NEW = {"CCrackWidthSP35Calculator", "CCrackWidthSP63Calculator", "modCrackWidthSupport",
       "CStabilitySP35Calculator", "CStabilitySP63Calculator", "modStabilityPlane"}
CHANGED = {"CCrackWidthCalculator", "CStabilityCalculator", "modTestBatchCalculation"}


def digest(path):
    """Отпечаток фактически проверяемых байтов файла."""
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def components(path):
    """Читает нативный экспорт, сохраняя строки каждого VBA-компонента."""
    text = path.read_text(encoding="utf-8-sig")
    pattern = r"(?ms)^={100}\nCOMPONENT: ([^\n]+)\nTYPE: (\d+)\nLINES: (\d+)\n={100}\n(.*?)(?=^={100}\nCOMPONENT: |\Z)"
    return {m[1]: (m[2], m[4]) for m in re.finditer(pattern, text)}


def normalized(text):
    """Учитывает изменение регистра идентификаторов редактором VBA, сохраняя строки."""
    tokens = []
    for part in re.findall(r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27]+', text):
        if part.startswith(('"', "'")):
            tokens.append(part)
        else:
            tokens.extend(re.findall(r"[A-Za-z_]\w*|\d+(?:\.\d*)?(?:[Ee][+-]?\d+)?|[^\s]", part.lower()))
    return "\n".join(tokens)


def source_equality(directory):
    """Проверяет состав проекта и единственную актуальную реализацию каждого класса."""
    old = components(directory / "OriginalNative.txt")
    new = components(directory / "NativeCandidate.txt")
    assert new.keys() - old.keys() == NEW
    assert not old.keys() - new.keys()
    changed = {n for n in old if normalized(old[n][1]) != normalized(new[n][1])}
    assert changed == CHANGED, changed
    files = {p.stem: p for folder in ("src", "tests") for p in (ROOT / folder).rglob("*") if p.suffix in (".cls", ".bas")}
    for name in CHANGED | NEW:
        body = re.search(r"(?ms)^Option Explicit.*", files[name].read_text(encoding="utf-8-sig"))[0]
        if name == "CStabilityCalculator":
            # Редактор VBA сохраняет этот предельный литерал с 15 значащими цифрами.
            # Именно такая запись уже находится в исходной рабочей книге; источник
            # и порядок проверки не менялись, что отдельно проверяет ExpressionPreservation.
            assert "1.79769313486231E+308" in old[name][1]
            body = body.replace("1.797693134862315E+308", "1.79769313486231E+308")
        if normalized(body) != normalized(new[name][1]):
            import difflib
            diff = list(difflib.unified_diff(normalized(body).splitlines(), normalized(new[name][1]).splitlines(), fromfile="source", tofile="native"))
            (directory / (name + "_SourceDifference.txt")).write_text("\n".join(diff), encoding="utf-8")
            raise AssertionError(name)
        assert new[name][0] == ("1" if files[name].suffix == ".bas" else "2")
    result = {"Passed": True, "NativeComponents": len(new), "AddedComponents": sorted(NEW),
              "ChangedComponents": sorted(changed), "UnrelatedComponentsUnchanged": True,
              "BuildImports": "Build-Workbook.ps1 imports src/Crack and src/Stability recursively.",
              "EditorLiteral": {"Source": "1.797693134862315E+308", "SavedByVBE": "1.79769313486231E+308", "SameInOriginalNativeCode": True}}
    (directory / "SourceEquality.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return result


def verify_fixture_delta(directory):
    """Допускает только исправление трёх недопустимых нулей в отдельной новой пробе."""
    import importlib.util
    spec = importlib.util.spec_from_file_location("normative_verifier", Path(__file__).with_name("Verify-NormativeCalculatorRefactor.py"))
    verifier = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(verifier)
    old, new = components(directory / "NativeLocalTested.txt"), components(directory / "NativeCandidate.txt")
    assert old.keys() == new.keys()
    changed = {n for n in old if normalized(old[n][1]) != normalized(new[n][1])}
    assert changed == {"modTestBatchCalculation"}, changed
    old_methods, new_methods = verifier.methods(old["modTestBatchCalculation"][1]), verifier.methods(new["modTestBatchCalculation"][1])
    assert old_methods.keys() == new_methods.keys()
    changed_methods = {k for k in old_methods if normalized(old_methods[k]) != normalized(new_methods[k])}
    key = ("RunNormativePartialBranchFailureTest", "Function")
    assert changed_methods == {key}, changed_methods
    corrected, count = re.subn(r"(?m)^([ \t]*)0#, 0#, (0\.7, 0\.15, 1\.5, 1#, 0\.7,)", r"\g<1>1#, 1#, \g<2>", old_methods[key])
    assert count == 3 and normalized(corrected) == normalized(new_methods[key])
    return {"OnlyChangedMethod": key[0], "Change": "Three probe calls use valid User eccentricities 1 mm instead of invalid zero.",
            "ProductionAlgorithmsAndMainDirectedTestMethodsUnchanged": True,
            "MainDirectedReportBytesDifferFromFinalProbeVersion": True}


def acceptance(directory, run_directory):
    """Требует полноту адресного плана и совпадение численных строк отчётов без допусков."""
    runs = json.loads((directory / run_directory / "DirectedRuns.json").read_text(encoding="utf-8-sig"))
    repeated = json.loads((directory / "Directed04/DirectedRuns.json").read_text(encoding="utf-8-sig"))
    if isinstance(repeated, dict):
        repeated = [repeated]
    assert len(repeated) == 1 and repeated[0]["Label"] == "Candidate" and repeated[0]["Mode"] == "Yes"
    assert not any(r["Label"] == "Candidate" and r["Mode"] == "Yes" for r in runs)
    runs += repeated
    assert {(r["Label"], r["Mode"]) for r in runs} == {(a, b) for a in ("Baseline", "Candidate") for b in ("No", "Yes")}
    assert len(runs) == 4 and all(r["Passed"] and r["Suites"] == 19 for r in runs)
    local_candidate = directory / "RC_Section_NDM.xlsm"
    candidate = directory / "FinalCandidate.xlsm"
    assert {r["WorkbookSHA256"] for r in runs if r["Label"] == "Candidate" and r["Mode"] == "No"} == {digest(local_candidate)}
    assert {r["WorkbookSHA256"] for r in runs if r["Label"] == "Candidate" and r["Mode"] == "Yes"} == {digest(candidate)}
    fixture_delta = verify_fixture_delta(directory)
    probes = json.loads((directory / "PartialBranch02/Acceptance.json").read_text(encoding="utf-8-sig"))
    assert len(probes) == 4 and {(p["Label"], p["Mode"]) for p in probes} == {(a, b) for a in ("Baseline", "Candidate") for b in ("No", "Yes")}
    assert all(p["Passed"] and p["Assertions"] == 11 for p in probes)
    assert {p["WorkbookSHA256"] for p in probes if p["Label"] == "Candidate"} == {digest(candidate)}
    for probe in probes:
        report = ROOT / probe["Report"]
        assert json.loads(Path(str(report) + ".ProcessClosure.json").read_text(encoding="utf-8-sig"))["Passed"]
    comparisons = []
    for mode in ("No", "Yes"):
        reports = []
        for label in ("Baseline", "Candidate"):
            entry = next(r for r in runs if r["Label"] == label and r["Mode"] == mode)
            report = ROOT / entry["Report"]
            text = report.read_text(encoding="utf-8-sig")
            assert not re.search(r"(?m)^FAIL[: ]|failed=[1-9]|RUNTIME ERROR|WATCHDOG_FAILURE", text)
            assert text.count("SUITE_FINISHED:") == 19
            assert text.count("NORMATIVE_CORE_GUARD_RESTORED: True") == 19
            assert json.loads(Path(str(report) + ".ProcessClosure.json").read_text(encoding="utf-8-sig"))["Passed"]
            lines = [line for line in text.splitlines() if line.startswith(("OK:", "TOTAL", "PROJECTION_", "STABILITY_CONFIG:"))]
            # Длительность — наблюдение производительности, а не физический результат.
            lines = [re.sub(r"; elapsedSec=[^;\r\n]+", "", line) for line in lines]
            reports.append(lines)
        if reports[0] != reports[1]:
            import difflib
            differences = list(difflib.unified_diff(reports[0], reports[1]))
            (directory / f"ReportDifference_{mode}.txt").write_text("\n".join(differences), encoding="utf-8")
            raise AssertionError(f"Different engineering assertions: {mode}")
        comparisons.append({"Mode": mode, "MatchingEngineeringLines": len(reports[0]), "ComparedWithoutNumericTolerance": True})
    proof = json.loads((directory / "ExpressionPreservation.json").read_text(encoding="utf-8"))
    assert proof["Status"] == "PASS" and proof["Methods"] == 239
    return {"Status": "DIRECTED_PASS", "SuitesPerVersionAndMode": 19, "Executions": 76,
            "ExpressionPreservationMethods": proof["Methods"], "Comparisons": comparisons,
            "FullAcceptanceRun": False, "ExcelStateRestoredAfterEachSuite": True, "RunDirectory": run_directory,
            "FinalProbeExecutions": 4, "FixtureCorrection": fixture_delta,
            "MainLocalCandidateSHA256": digest(local_candidate), "FinalPrivateCandidateSHA256": digest(candidate)}


def package(directory):
    """Готовит книгу с единственной заменённой частью xl/vbaProject.bin."""
    source = ROOT / "workbook/output/RC_Section_NDM.xlsm"
    assert digest(source) == EXPECTED_ORIGINAL, "Production workbook changed; preserve the latest user copy before packaging."
    accepted = directory / "Accepted.xlsm"
    assert not accepted.exists(), "Accepted package already exists."
    with ZipFile(source) as original, ZipFile(directory / "FinalCandidate.xlsm") as compiled, ZipFile(accepted, "w") as output:
        for item in original.infolist():
            output.writestr(item, compiled.read(item.filename) if item.filename == "xl/vbaProject.bin" else original.read(item.filename))
    with ZipFile(source) as original, ZipFile(accepted) as output:
        assert original.namelist() == output.namelist()
        changed = [n for n in original.namelist() if original.read(n) != output.read(n)]
        assert changed == ["xl/vbaProject.bin"], changed
    export = (directory / "NativeCandidate.txt").read_text(encoding="utf-8-sig")
    export = re.sub(r"(?m)^Date: [^\n]*\n", "", export)
    export = re.sub(r"(?m)^Workbook: [^\n]*", "Workbook: RC_Section_NDM.xlsm", export, count=1)
    (directory / "VBA_All_Code.txt").write_text(export, encoding="utf-8-sig", newline="\r\n")
    return {"AcceptedSHA256": digest(accepted), "OriginalSHA256": EXPECTED_ORIGINAL,
            "AllNonVBAPartsByteIdentical": True, "ChangedZIPParts": changed,
            "VBAExportSHA256": digest(directory / "VBA_All_Code.txt")}


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--directory", type=Path, required=True)
    parser.add_argument("--sources-only", action="store_true")
    parser.add_argument("--run-directory", default="Directed03")
    args = parser.parse_args()
    directory = (ROOT / args.directory).resolve()
    result = {"Sources": source_equality(directory)}
    if not args.sources_only:
        result["Acceptance"] = acceptance(directory, args.run_directory)
        result["Package"] = package(directory)
        (directory / "AcceptanceAndPackage.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False))
