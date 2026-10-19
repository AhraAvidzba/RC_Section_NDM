"""Независимая XML-сверка данных книги при добавлении концевых отступов Config."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def column_number(text: str) -> int:
    result = 0
    for letter in text:
        result = result * 26 + ord(letter) - 64
    return result


def read_book(path: Path) -> dict:
    with ZipFile(path) as archive:
        shared = ["".join(item.itertext()) for item in ET.fromstring(archive.read("xl/sharedStrings.xml"))]
        book = ET.fromstring(archive.read("xl/workbook.xml"))
        relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {item.attrib["Id"]: item.attrib["Target"] for item in relations}
        sheets = {}
        for item in book.findall("s:sheets/s:sheet", NS):
            rid = item.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = targets[rid]
            xml = ET.fromstring(archive.read(target.lstrip("/") if target.startswith("/") else "xl/" + target))
            cells = {}
            for cell in xml.findall(".//s:sheetData/s:row/s:c", NS):
                match = re.fullmatch(r"([A-Z]+)(\d+)", cell.attrib["r"])
                value = cell.find("s:v", NS)
                formula = cell.find("s:f", NS)
                text = value.text if value is not None else None
                if cell.attrib.get("t") == "s": text = shared[int(text)]
                cells[(int(match[2]), column_number(match[1]))] = (formula.text if formula is not None else None, text)
            sheets[item.attrib["name"]] = cells
        names = {item.attrib["name"]: item.text for item in book.findall("s:definedNames/s:definedName", NS)}
        return {"names": names, "sheets": sheets}


def table(book: dict, name: str) -> list[list[tuple]]:
    reference = book["names"][name]
    sheet, address = reference.rsplit("!", 1)
    sheet = sheet.strip("'")
    match = re.fullmatch(r"\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)", address)
    if not match: raise ValueError("Expected a rectangular named range: " + reference)
    c1, r1, c2, r2 = column_number(match[1]), int(match[2]), column_number(match[3]), int(match[4])
    return [[book["sheets"][sheet].get((r, c), (None, None)) for c in range(c1, c2+1)] for r in range(r1, r2+1)]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("before")
    parser.add_argument("after")
    parser.add_argument("--report", required=True)
    args = parser.parse_args()
    paths = [Path(args.before), Path(args.after)]
    before, after = [read_book(path) for path in paths]
    errors = []
    added = sorted(set(after["names"]) - set(before["names"]))
    removed = sorted(set(before["names"]) - set(after["names"]))
    if added != ["rngSectionRecipeTypes"] or removed: errors.append("Unexpected named range delta")
    checks = 0
    for name, reference in before["names"].items():
        if name.startswith("_xlnm."): continue
        if after["names"].get(name) != reference: errors.append("Reference changed: " + name)
        if "!" not in reference or not re.search(r"\$[A-Z]+\$\d+:\$[A-Z]+\$\d+$", reference): continue
        old, new = table(before, name), table(after, name)
        main = next((r for r, values in enumerate(old) if len(values)>1 and values[1][1] == "as"), None)
        face_count = {"rngPointedRectGeometry":4,"rngHollowRectangleGeometry":8}.get(name)
        for r, values in enumerate(old):
            for c, value in enumerate(values):
                expected = value
                if face_count and r == main and 4 <= c <= 7:
                    expected = (None, ["StartSetback","EndSetback","Ед.","Комментарий"][c-4])
                elif face_count and main < r <= main+face_count and 4 <= c <= 7:
                    expected = (None,"0") if c < 6 else old[r][c-2]
                elif name == "rngRectSetGeometry" and r == main and c in [4,5]:
                    expected = (None, ["StartSetback","EndSetback"][c-4])
                elif name == "rngCircleGeometry" and values[0][1] == "Rebar.Count" and c == 3:
                    checks += 1
                    if not new[r][c][1].startswith(value[1]) or "ClosedUniform" not in new[r][c][1]: errors.append("Closed circle explanation is not an append")
                    continue
                checks += 1
                if new[r][c] != expected: errors.append(f"{name}[{r+1},{c+1}]: {expected!r} -> {new[r][c]!r}")
    report = {"status":"FAIL" if errors else "PASS", "checkedCells":checks, "addedNames":added,
              "removedNames":removed, "errors":errors, "sha256":[hashlib.sha256(path.read_bytes()).hexdigest() for path in paths]}
    Path(args.report).write_text(json.dumps(report,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    print(json.dumps({key:report[key] for key in ["status","checkedCells","addedNames","removedNames"]},ensure_ascii=False))
    raise SystemExit(bool(errors))


if __name__ == "__main__": main()
