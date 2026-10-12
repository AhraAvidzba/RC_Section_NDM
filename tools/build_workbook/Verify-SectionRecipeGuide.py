"""Независимая XML-приёмка Справки, старых листов, формул и форматирования."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
from functools import lru_cache
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def canonical(element):
    if element is None:
        return None
    return element.tag, tuple(sorted(element.attrib.items())), element.text, tuple(canonical(child) for child in element)


def read(path: Path):
    with ZipFile(path) as archive:
        strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            strings = ["".join(item.itertext()) for item in ET.fromstring(archive.read("xl/sharedStrings.xml"))]
        styles = ET.fromstring(archive.read("xl/styles.xml"))
        containers = {name: list(styles.find(S + name)) for name in ("fonts", "fills", "borders", "cellXfs", "cellStyleXfs")}
        @lru_cache(maxsize=None)
        def style(index):
            xf = containers["cellXfs"][index]
            attributes = dict(xf.attrib)
            refs = []
            for attr, group in (("fontId", "fonts"), ("fillId", "fills"), ("borderId", "borders"), ("xfId", "cellStyleXfs")):
                refs.append(canonical(containers[group][int(attributes.pop(attr, "0"))]))
            return tuple(sorted(attributes.items())), tuple(refs), tuple(canonical(child) for child in xf)
        book = ET.fromstring(archive.read("xl/workbook.xml"))
        relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {item.attrib["Id"]: item.attrib["Target"] for item in relations}
        sheets = {}
        for item in book.find(S + "sheets"):
            target = targets[item.attrib[R + "id"]]
            document = ET.fromstring(archive.read(target.lstrip("/") if target.startswith("/") else "xl/" + target))
            cells, rows = {}, {}
            for row in document.find(S + "sheetData"):
                number = int(row.attrib["r"])
                row_attrs = dict(row.attrib)
                if "s" in row_attrs:
                    row_attrs["s"] = style(int(row_attrs["s"]))
                rows[number] = row_attrs
                for cell in row:
                    address = cell.attrib["r"]
                    value = cell.find(S + "v")
                    text = value.text if value is not None else None
                    if cell.attrib.get("t") == "s":
                        text = strings[int(text)]
                    elif cell.attrib.get("t") == "inlineStr":
                        text = "".join(cell.find(S + "is").itertext())
                    formula = cell.find(S + "f")
                    attrs = dict(cell.attrib)
                    attrs.pop("s", None); attrs.pop("t", None)
                    cells[address] = {"formula": canonical(formula), "value": text,
                                      "style": style(int(cell.attrib.get("s", "0"))), "attrs": attrs}
            sheets[item.attrib["name"]] = {"cells": cells, "rows": rows, "document": document}
        names = {item.attrib["name"]: item.text for item in book.find(S + "definedNames")}
        drawings = {name: canonical(ET.fromstring(archive.read(name))) for name in archive.namelist()
                    if name.startswith("xl/drawings/") and name.endswith(".xml")}
        return {"sheets": sheets, "names": names, "drawings": drawings}


def row_of(address):
    return int(re.search(r"[0-9]+$", address).group())


def verify(before: Path, after: Path, layout: Path, existing_guide: bool = False):
    old, new = read(before), read(after)
    content = json.loads(layout.read_text(encoding="utf-8-sig"))
    name = new["names"].get("rngSectionRecipeGuide", "")
    match = re.fullmatch(r"'?Справка'?!\$A\$(\d+):\$F\$(\d+)", name)
    if not match:
        raise ValueError("Missing guide owner range: " + name)
    start, end = map(int, match.groups())
    errors, checks, cached = [], 0, []
    added = sorted(set(new["names"]) - set(old["names"]))
    expected_added = [] if existing_guide else ["rngSectionRecipeGuide"]
    if existing_guide and "rngSectionRecipeGuide" not in old["names"]:
        errors.append("Existing-guide reference does not contain rngSectionRecipeGuide")
    if added != expected_added:
        errors.append("Unexpected names added: " + repr(added))
    for key, value in old["names"].items():
        if new["names"].get(key) != value:
            errors.append("Old named range changed: " + key)
    for sheet, data in old["sheets"].items():
        actual = new["sheets"][sheet]
        populated_rows = {row_of(address) for address, cell in data["cells"].items() if cell["value"] or cell["formula"]}
        for address, expected in data["cells"].items():
            value = actual["cells"].get(address)
            checks += 1
            if value != expected:
                if expected["formula"] is not None and value and all(value[k] == expected[k] for k in ("formula", "style", "attrs")):
                    cached.append(sheet + "!" + address)
                else:
                    errors.append("Old cell changed: " + sheet + "!" + address)
        for row, attrs in data["rows"].items():
            if sheet == "Справка" and row >= start:
                # Форматированная пустота UsedRange не является содержанием.
                if row not in populated_rows:
                    continue
            if actual["rows"].get(row) != attrs:
                errors.append("Old row metadata changed: " + sheet + "!" + str(row))
        for tag in ("cols", "mergeCells", "hyperlinks"):
            previous = data["document"].find(S + tag)
            following = actual["document"].find(S + tag)
            if tag == "cols":
                if canonical(previous) != canonical(following): errors.append("Column widths/styles changed: " + sheet)
            elif previous is not None:
                values = {canonical(item) for item in following} if following is not None else set()
                for item in previous:
                    if canonical(item) not in values: errors.append("Old " + tag + " changed: " + sheet + ":" + str(item.attrib))
    for drawing, value in old["drawings"].items():
        if new["drawings"].get(drawing) != value:
            errors.append("Old drawing changed: " + drawing)
    help_cells = new["sheets"]["Справка"]["cells"]
    if end - start + 1 != len(content["rows"]): errors.append("Guide row count mismatch")
    for index, row in enumerate(content["rows"]):
        fields = row.get("cells", [row["text"]])
        for column, field in enumerate(fields):
            cell = help_cells.get(chr(65+column) + str(start + index), {})
            if (cell.get("value") or "") != field or cell.get("formula") is not None:
                errors.append("Guide source text mismatch: " + str(index) + ":" + str(column))
    values = [cell for address, cell in help_cells.items() if row_of(address) > end and (cell["value"] or cell["formula"])]
    if values: errors.append("Content follows guide")
    return {"schema": 1, "status": "XML_PRESERVATION_PASS" if not errors else "FAIL",
            "beforeSHA256": hashlib.sha256(before.read_bytes()).hexdigest(), "afterSHA256": hashlib.sha256(after.read_bytes()).hexdigest(),
            "oldCellsChecked": checks, "oldNamedRanges": len(old["names"]), "oldDrawings": len(old["drawings"]),
            "addedNames": added, "formulaCacheUpdates": cached, "guideStart": start, "guideEnd": end,
            "guideRows": len(content["rows"]), "members": content["members"], "apiStatus": content["apiStatus"], "errors": errors}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("before", type=Path); parser.add_argument("after", type=Path)
    parser.add_argument("--existing-guide", action="store_true", help="Strictly compare an already installed guide: no new names expected")
    parser.add_argument("--layout", type=Path, default=Path("docs/geometry_refactor/SECTION_RECIPE_GUIDE_LAYOUT.json"))
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    result = verify(args.before, args.after, args.layout, args.existing_guide)
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({key: result[key] for key in ("status", "oldCellsChecked", "guideRows", "members", "apiStatus", "errors")}, ensure_ascii=False))
    if result["errors"]: raise SystemExit(1)
