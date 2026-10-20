"""Read-only acceptance of preserved inputs and the saved dependent-row layout."""
from pathlib import Path
from zipfile import ZipFile
from decimal import Decimal
import posixpath
import argparse
import json
import re
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
D = ROOT / "docs/regression/DependentRowOffsets_2026-10-17"
N = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def column(n):
    s = ""
    while n:
        n, c = divmod(n - 1, 26)
        s = chr(65 + c) + s
    return s


def point(address):
    m = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)", address)
    c = 0
    for ch in m[1]:
        c = c * 26 + ord(ch) - 64
    return int(m[2]), c


class Book:
    def __init__(self, path):
        self.path = path
        self.zip = ZipFile(path)
        self.root = ET.fromstring(self.zip.read("xl/workbook.xml"))
        rels = ET.fromstring(self.zip.read("xl/_rels/workbook.xml.rels"))
        targets = {x.attrib["Id"]: x.attrib["Target"].lstrip("/") for x in rels}
        strings = ET.fromstring(self.zip.read("xl/sharedStrings.xml"))
        texts = ["".join(x.itertext()) for x in strings]
        self.sheets = {}
        for x in self.root.find("s:sheets", N):
            rid = x.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = targets[rid]
            if not target.startswith("xl/"):
                target = "xl/" + target
            sheet = ET.fromstring(self.zip.read(target))
            cells = {}
            for c in sheet.findall(".//s:sheetData/s:row/s:c", N):
                f = c.find("s:f", N)
                v = c.find("s:v", N)
                value = v.text if v is not None and v.text else ""
                if c.attrib.get("t") == "s":
                    value = texts[int(value)]
                elif c.attrib.get("t") == "inlineStr":
                    value = "".join(c.find("s:is", N).itertext())
                elif c.attrib.get("t", "n") == "n" and value:
                    value = Decimal(value)
                formula = f.text or "" if f is not None else ""
                cells[c.attrib["r"]] = {"value": value, "formula": formula, "style": int(c.attrib.get("s", 0))}
            comments = {}
            relfile = posixpath.join(posixpath.dirname(target), "_rels", posixpath.basename(target) + ".rels")
            if relfile in self.zip.namelist():
                for rel in ET.fromstring(self.zip.read(relfile)):
                    if rel.attrib["Type"].endswith("/comments"):
                        note_path = posixpath.normpath(posixpath.join(posixpath.dirname(target), rel.attrib["Target"]))
                        for note in ET.fromstring(self.zip.read(note_path)).findall(".//s:comment", N):
                            comments[note.attrib["ref"]] = "".join(note.find("s:text", N).itertext())
            self.sheets[x.attrib["name"]] = {"xml": sheet, "cells": cells, "comments": comments}
        self.names = {x.attrib["name"]: x.text for x in self.root.find("s:definedNames", N)}
        self.styles = ET.fromstring(self.zip.read("xl/styles.xml"))

    def bounds(self, name):
        return tuple(point(x) for x in self.names[name].split("!")[1].split(":"))

    def cell(self, sheet, row, col):
        return self.sheets[sheet]["cells"].get(column(col) + str(row), {"value": "", "formula": "", "style": 0})


def equivalent(a, b):
    if a["formula"] or b["formula"]:
        return a["formula"] == b["formula"]
    return a["value"] == b["value"]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--target", default="Accepted.xlsm")
    args = parser.parse_args()
    a, b = Book(D / "Baseline.xlsm"), Book(D / args.target)
    failures, comparisons, tables = [], 0, []
    geometry = ["rngCircleGeometry", "rngRectanglesGeometry", "rngPointedSectionGeometry", "rngHollowRectangleGeometry"]
    old_areas = [a.bounds(n) for n in geometry]
    for name in geometry:
        (ar, ac), (ae, ax) = a.bounds(name)
        (br, bc), (be, bx) = b.bounds(name)
        if bx - bc != 10:
            failures.append(f"{name}: not eleven columns")
        header = next((r for r in range(ar, ae + 1) if a.cell("Config", r, ac + 1)["value"] == "d2"), None)
        for row in range(ar, ae + 1):
            for col in range(ac, ax + 1):
                source = a.cell("Config", row, col)
                if name == "rngCircleGeometry":
                    key = a.cell("Config", row, ac)["value"]
                    main_cols = {"Rebar.AxisDistance": 1, "Rebar.Count": 3, "Rebar.Diameter": 2}
                    extra_cols = {"Rebar.Diameter2": 1, "Rebar.Bind2row": 3, "Rebar.Loc2row": 4,
                                  "Rebar.Diameter3": 5, "Rebar.Bind3row": 7, "Rebar.Loc3row": 8}
                    if key == "Circle.Diameter" and col > ac:
                        target = b.cell("Config", br + 2, bc + col - ac)
                    elif key in main_cols and col == ac + 1:
                        target = b.cell("Config", br + 6, bc + main_cols[key])
                    elif key in main_cols and col == ac + 3:
                        note = b.sheets["Config"]["comments"].get(column(bc + main_cols[key]) + str(br + 6), "")
                        comparisons += 1
                        if note != source["value"]:
                            failures.append(f"Circle original explanation missing: {key}")
                        continue
                    elif key in extra_cols and col == ac + 1:
                        target = b.cell("Config", br + 10, bc + extra_cols[key])
                    elif key == "Rebar.Diameter2" and col == ac + 2:
                        target = b.cell("Config", br + 10, bc + 9)
                    else:
                        continue
                elif header is not None and row >= header:
                    if row == header:
                        continue
                    mapping = [0, 1, 4, 3, 5, 8, 7, 9, 10]
                    target = b.cell("Config", br + row - ar, bc + mapping[col - ac])
                else:
                    target = b.cell("Config", br + row - ar, bc + col - ac)
                comparisons += 1
                if not equivalent(source, target):
                    failures.append(f"{name}: input changed at {column(col)}{row}: {source} -> {target}")
        tables.append({"Name": name, "Range": b.names[name]})
    # Every meaningful cell outside the four transformed blocks and Help keeps its value/formula.
    for sheet, record in a.sheets.items():
        if sheet == "Справка":
            continue
        for address, cell in record["cells"].items():
            if not cell["formula"] and cell["value"] == "":
                continue
            row, col = point(address)
            if sheet == "Config" and any(r - 1 <= row <= e and c <= col <= x for (r, c), (e, x) in old_areas):
                continue
            comparisons += 1
            if not equivalent(cell, b.cell(sheet, row, col)):
                failures.append(f"Outside geometry changed: {sheet}!{address}")
    for name, refers in a.names.items():
        if name in geometry or "Print_Area" in name or name == "rngSectionRecipeGuide":
            continue
        if b.names.get(name) != refers:
            failures.append(f"Name changed: {name}")
    # Saved geometry may have dashed or absent borders; solid lines are forbidden.
    xfs = list(b.styles.find("s:cellXfs", N))
    borders = list(b.styles.find("s:borders", N))
    for name in geometry:
        (r, c), (e, x) = b.bounds(name)
        for row in range(r, e + 1):
            for col in range(c, x + 1):
                border = borders[int(xfs[b.cell("Config", row, col)["style"]].attrib["borderId"])]
                for side in ("left", "right", "top", "bottom"):
                    edge = border.find("s:" + side, N)
                    if edge is not None and edge.attrib.get("style") not in (None, "dashed"):
                        failures.append(f"Solid geometry border: {name}/{column(col)}{row}/{side}")
    guide = b.bounds("rngSectionRecipeGuide")
    if guide[1][0] - guide[0][0] + 1 != 61:
        failures.append("Recipe guide range differs from its 61-row content")
    result = {"Passed": not failures, "Comparisons": comparisons, "Tables": tables, "Failures": failures,
              "WorkbookCalculationBefore": ET.tostring(a.root.find("s:calcPr", N), encoding="unicode"),
              "WorkbookCalculationAfter": ET.tostring(b.root.find("s:calcPr", N), encoding="unicode")}
    (D / "SavedArtifactAudit.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(bool(failures))


if __name__ == "__main__":
    main()
