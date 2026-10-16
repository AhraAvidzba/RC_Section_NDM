"""Приёмка обязательных номеров сочетаний в таблице нагрузок устойчивости."""
import argparse
import hashlib
import importlib.util
import json
import re
from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
D = ROOT / "docs/regression/ExcelGuardGuide_2026-10-13/StabilityDurationMatch_2026-10-16"


def module(name, file):
    spec = importlib.util.spec_from_file_location(name, Path(__file__).with_name(file))
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


presentation = module("duration_presentation", "Verify-GuidePresentation.py")
reader = presentation.reader


def receipt(name):
    return json.loads((D / name).read_text(encoding="utf-8-sig"))


def payload(name):
    path = D / name
    if path.exists():
        return path.read_bytes()
    manifest = receipt("Archive.json")
    assert sha(D / "Evidence.zip") == manifest["ArchiveSHA256"]
    with ZipFile(D / "Evidence.zip") as archive:
        data = archive.read(name)
        assert hashlib.sha256(data).hexdigest().upper() == manifest["Entries"][name]
        return data


def book_data(name):
    path = D / name
    return path if path.exists() else BytesIO(payload(name.replace(".xlsm", ".data.zip")))


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def normalized(body, comments=True):
    parts = re.findall(r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27\n]+|\n', body)
    return "".join(part if part.startswith('"') or (comments and part.startswith("'")) else
                   "" if part.startswith("'") else part.lower() for part in parts).strip()


def read_components(name):
    text = payload(name).decode("utf-8-sig").replace("\r\n", "\n")
    pattern = r"(?ms)^={100}\nCOMPONENT: ([^\n]+)\nTYPE: (\d+)\nLINES: (\d+)\n={100}\n(.*?)(?=^={100}\nCOMPONENT: |\Z)"
    return {m[1]: (m[2], m[3], m[4]) for m in re.finditer(pattern, text)}


def tests():
    before = receipt("BeforeFix/Run.json")
    assert not before["Passed"] and before["ProductionUnchanged"]
    old = payload("BeforeFix/modTestProfileConfig.RunStabilityDurationMatchingTests.txt").decode("utf-8-sig")
    assert "passed=234; failed=80; cases=24" in old
    runs = []
    for stage, macro, cases in (("Matching", "RunStabilityDurationMatchingTests", 24),
                                 ("DurationRegression", "RunAudit03DurationConfigTests", 637)):
        run = receipt(stage + "/Run.json")
        assert run["Passed"] and run["ProductionUnchanged"] and len(run["Runs"]) == 1
        assert run["Runs"][0]["Macro"] == "modTestProfileConfig." + macro
        text = payload(stage + "/modTestProfileConfig." + macro + ".txt").decode("utf-8-sig")
        assert not re.search(r"(?m)^FAIL:|RUNTIME ERROR:", text)
        found = re.findall(r"TOTAL_AUDIT03_DURATION_CONFIG: passed=(\d+); failed=(\d+); cases=(\d+)", text)
        assert len(found) == 1
        passed, failed, count = map(int, found[0])
        assert passed > 0 and failed == 0 and (cases is None or count == cases)
        assert len(re.findall(r"(?m)^RUN: durationMatching\.", text)) == 24
        runs.append(dict(Stage=stage, Passed=passed, Failed=failed, Cases=count,
                         Seconds=run["Runs"][0]["Seconds"], WorkbookSHA256=run["SourceSHA256"]))
    assert runs[0]["WorkbookSHA256"] == runs[1]["WorkbookSHA256"]
    return runs


def scope():
    baseline, candidate = read_components("BaselineNativeVBA.txt"), read_components("CandidateNativeVBA.txt")
    assert baseline.keys() == candidate.keys() and len(candidate) == 158
    changed = sorted(name for name in baseline if normalized(baseline[name][2]) != normalized(candidate[name][2]))
    assert changed == ["CBatchSectionCalculator", "modTestProfileConfig", "modWorkbookCalculation"], changed
    executable = sorted(name for name in baseline if normalized(baseline[name][2], False) != normalized(candidate[name][2], False))
    assert executable == ["modTestProfileConfig", "modWorkbookCalculation"], executable
    assert receipt("CandidateSourceEquality.json")["Passed"]
    return changed


def new_help():
    source = (ROOT / "tools/build_workbook/SettingsCatalog.ps1").read_text(encoding="utf-8-sig")
    found = re.findall(r'^\s*"(При включённой устойчивости отсутствие нужного CombinationID[^\n]+)",$', source, re.M)
    assert len(found) == 1
    return found[0]


def package():
    runs = tests(); changed = scope()
    assert sha(D / "Candidate/Matching.xlsm") == runs[0]["WorkbookSHA256"]
    initial, user = read_components("BaselineNativeVBA.txt"), read_components("UserNativeVBA.txt")
    assert initial.keys() == user.keys()
    assert all(normalized(initial[n][2]) == normalized(user[n][2]) for n in initial), "User VBA changed"
    user_book = D / "UserBaseline.xlsm"
    old = reader.read(user_book)
    matches = [(address, cell["value"]) for address, cell in old["sheets"]["Справка"]["cells"].items()
               if (cell["value"] or "").startswith("Пустое число или отсутствие строки сочетания")]
    assert len(matches) == 1, matches
    address, old_text = matches[0]
    new_text = new_help()
    accepted = D / "Accepted.xlsm"
    assert not accepted.exists()
    with ZipFile(user_book) as a, ZipFile(D / "Candidate/Matching.xlsm") as b, ZipFile(accepted, "w") as out:
        old_strings = a.read("xl/sharedStrings.xml")
        needle, replacement = escape(old_text).encode("utf-8"), escape(new_text).encode("utf-8")
        assert old_strings.count(needle) == 1
        for info in a.infolist():
            content = a.read(info.filename)
            if info.filename == "xl/vbaProject.bin":
                content = b.read(info.filename)
            elif info.filename == "xl/sharedStrings.xml":
                content = content.replace(needle, replacement)
            out.writestr(info, content)
    with ZipFile(user_book) as a, ZipFile(accepted) as b:
        assert a.namelist() == b.namelist()
        parts = sorted(n for n in a.namelist() if a.read(n) != b.read(n))
        assert parts == ["xl/sharedStrings.xml", "xl/vbaProject.bin"]
    result = dict(Passed=True, ChangedNativeComponents=changed, NativeComponents=158,
                  HelpCell=address, OldHelp=old_text, NewHelp=new_text, ChangedZIPParts=parts,
                  AcceptedSHA256=sha(accepted), UserBaselineSHA256=sha(user_book))
    (D / "Package.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({k: v for k, v in result.items() if k not in ("OldHelp", "NewHelp")}, ensure_ascii=False))


def preservation():
    package_data = receipt("Package.json")
    old, new = reader.read(book_data("UserBaseline.xlsm")), reader.read(book_data("Accepted.xlsm"))
    assert old["names"] == new["names"] and old["drawings"] == new["drawings"]
    assert old["sheets"].keys() == new["sheets"].keys()
    checked = 0
    for name, sheet in old["sheets"].items():
        actual = new["sheets"][name]
        assert sheet["rows"] == actual["rows"]
        assert reader.canonical(sheet["document"]) == reader.canonical(actual["document"])
        assert sheet["cells"].keys() == actual["cells"].keys()
        for address, before in sheet["cells"].items():
            after = actual["cells"][address]
            checked += 1
            if name == "Справка" and address == package_data["HelpCell"]:
                assert before["value"] == package_data["OldHelp"] and after["value"] == package_data["NewHelp"]
                assert all(before[k] == after[k] for k in ("formula", "style", "attrs"))
            else:
                assert before == after, (name, address)
    return dict(Cells=checked, NamedRanges=len(old["names"]), OnlyHelpCellChanged=package_data["HelpCell"])


def accept(published):
    runs = tests(); changed = scope(); preserved = preservation()
    accepted = read_components("AcceptedNativeVBA.txt")
    assert accepted == read_components("CandidateNativeVBA.txt")
    assert receipt("AcceptedSourceEquality.json")["Passed"]
    marker = re.compile(r"NDM\.RECIPE\.GUIDE/\d|\b[0-9a-f]{64}\b", re.I)
    with ZipFile(book_data("Accepted.xlsm")) as archive:
        assert all(not marker.search(archive.read(n).decode("utf-8-sig")) for n in archive.namelist() if n.endswith((".xml", ".rels")))
    assert not marker.search(payload("AcceptedNativeVBA.txt").decode("utf-8-sig"))
    if published:
        assert sha(ROOT / "workbook/output/RC_Section_NDM.xlsm") == receipt("Package.json")["AcceptedSHA256"]
        assert receipt("PublishedSourceEquality.json")["Passed"]
        assert read_components("PublishedNativeVBA.txt") == accepted
        assert (ROOT / "workbook/output/VBA_All_Code.txt").read_bytes() == payload("PublishedNativeVBA.txt")
        assert sha(ROOT / "workbook/output/RC_Section_NDM_execution_report.txt") == receipt("UserSnapshot.json")["Hashes"]["workbook/output/RC_Section_NDM_execution_report.txt"]
    result = dict(Passed=True, Runs=runs, ChangedNativeComponents=changed, Preserved=preserved, PublishedVerified=published)
    (D / "Acceptance.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--package", action="store_true")
    parser.add_argument("--published", action="store_true")
    args = parser.parse_args()
    package() if args.package else accept(args.published)
