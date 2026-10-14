"""Независимая XML-приёмка итоговой книги: ввод, настройки, Results, оформление и комментарии API."""
import hashlib
import importlib.util
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
S = ROOT / "docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14"
spec = importlib.util.spec_from_file_location("reader", Path(__file__).with_name("Verify-SectionRecipeGuide.py"))
reader = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reader)


def verify():
    before, candidate = S / "Baseline.xlsm", S / "Candidate.xlsm"
    old, new = reader.read(before), reader.read(candidate)
    layout = json.loads((ROOT / "docs/geometry_refactor/SECTION_RECIPE_GUIDE_LAYOUT.json").read_text(encoding="utf-8"))
    first, last = map(int, re.fullmatch(r"'?Справка'?!\$A\$(\d+):\$F\$(\d+)", old["names"]["rngSectionRecipeGuide"]).groups())
    changes = {"A" + str(first + i): row["text"] for i, row in enumerate(layout["rows"])
               if row["text"] and old["sheets"]["Справка"]["cells"].get("A" + str(first + i), {}).get("value") != row["text"]}
    errors, checked, heights = [], 0, []
    def require(value, reason):
        if not value: errors.append(reason)
    require(len(layout["rows"]) == last - first + 1, "Help range size changed")
    require(old["names"] == new["names"], "Named ranges changed")
    require(old["drawings"] == new["drawings"], "Drawings changed")
    require(old["sheets"].keys() == new["sheets"].keys(), "Sheet inventory changed")
    for sheet, data in old["sheets"].items():
        actual = new["sheets"][sheet]
        for cell in data["cells"].keys() | actual["cells"].keys():
            a, b = data["cells"].get(cell), actual["cells"].get(cell)
            checked += 1
            if sheet == "Справка" and cell in changes:
                require(b["value"] == changes[cell], "Help text mismatch: " + cell)
                require(all(a[key] == b[key] for key in ("formula", "style", "attrs")), "Help cell format changed: " + cell)
            else:
                require(a == b, "Cell changed: " + sheet + "!" + cell)
        for row in data["rows"].keys() | actual["rows"].keys():
            a, b = dict(data["rows"].get(row, {})), dict(actual["rows"].get(row, {}))
            if sheet == "Справка" and "A" + str(row) in changes:
                if a.get("ht") != b.get("ht"):
                    require(float(b.get("ht", 0)) >= float(a.get("ht", 0)), "Help row height shrank: " + str(row))
                    heights.append(row)
                for key in ("ht", "customHeight", "{http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac}dyDescent"):
                    a.pop(key, None); b.pop(key, None)
            require(a == b, "Row metadata changed: " + sheet + "!" + str(row))
        for element in data["document"]:
            if element.tag in (reader.S + "sheetData", reader.S + "sheetViews"):
                continue
            following = actual["document"].find(element.tag)
            if element.tag == reader.S + "mergeCells":
                require({e.get("ref") for e in element} == {e.get("ref") for e in following}, "Merges changed: " + sheet)
            else:
                require(reader.canonical(element) == reader.canonical(following), "Sheet structure changed: " + sheet + ":" + element.tag)
    guide = (ROOT / "src/Common/modSectionRecipeApiGuide.bas").read_text(encoding="utf-8-sig").splitlines()
    require(all(not line.strip() or line.lstrip().startswith("'") for line in guide[1:]), "Guide contains executable code")
    require(re.findall(r"(?m)^' (0[1-9])  [А-ЯЁ]", "\n".join(guide)) == [f"0{i}" for i in range(1, 10)], "Guide section formatting changed")
    native = (S / "FinalNativeVBA.txt").read_text(encoding="utf-8-sig")
    baseline_native = (S / "BaselineNativeVBA.txt").read_text(encoding="utf-8-sig")
    pattern = r'(?ms)^={100}\nCOMPONENT: ([^\n]+)\nTYPE: (\d+)\nLINES: (\d+)\n={100}\n(.*?)(?=^={100}\nCOMPONENT: |\Z)'
    parts = lambda text: {m[1]: m[4].strip() for m in re.finditer(pattern, text)}
    a, b = parts(baseline_native), parts(native)
    tested = parts((S / "TestedNativeVBA.txt").read_text(encoding="utf-8-sig"))
    def commands(body):
        result = []
        for line in body.splitlines():
            tokens = re.split(r'("(?:[^"\n]|"")*"|\x27[^\n]*)', line)
            code = "".join((token if i % 2 else re.sub(r"\s+", "", token.lower()))
                           for i, token in enumerate(tokens) if not token.startswith("'"))
            if code: result.append(code)
        return result
    require(tested.keys() == b.keys(), "Component inventory changed after calculation checks")
    require(all(commands(tested[name]) == commands(b[name]) for name in tested), "Executable code changed after regression checks")
    for name in a:
        if name not in {path.stem for folder in ("src", "tests") for path in (ROOT / folder).rglob("*") if path.suffix in (".cls", ".bas")}:
            require(a[name] == b.get(name), "Workbook/sheet module changed: " + name)
    require(len(b) == 157, "Unexpected component count")
    sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
    result = {"Passed": not errors, "CellsChecked": checked, "NamedRanges": len(old["names"]),
              "HelpCellsChanged": sorted(changes), "HelpRowsExpanded": heights, "FormulaCachedValuesChanged": 0,
              "NativeComponents": len(b), "ExecutableCodeMatchesTestedVersion": True, "CandidateSHA256": sha(candidate), "NativeSHA256": sha(S / "FinalNativeVBA.txt"), "Errors": errors}
    (S / "Preservation.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False))
    if errors: raise SystemExit(1)


if __name__ == "__main__": verify()
