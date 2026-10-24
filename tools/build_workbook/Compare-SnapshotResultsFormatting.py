"""ДЛЯ ТЕСТОВ: независимое сравнение оформления Results до и после оптимизации IO."""
from __future__ import annotations

import argparse
import hashlib
import json
from functools import lru_cache
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def canonical(node):
    """Канонизирует оформление без зависимости от порядка XML-атрибутов."""
    if node is None:
        return None
    return node.tag, tuple(sorted(node.attrib.items())), node.text, tuple(map(canonical, node))


def read(path: Path) -> dict:
    """Разрешает ссылки на стили и числовые форматы, сравнивая их смысл, а не номера."""
    with ZipFile(path) as archive:
        styles = ET.fromstring(archive.read("xl/styles.xml"))
        groups = {key: list(styles.find(S + key)) for key in ("fonts", "fills", "borders", "cellXfs", "cellStyleXfs")}
        number_formats = {node.attrib["numFmtId"]: node.attrib["formatCode"] for node in styles.findall(S + "numFmts/" + S + "numFmt")}

        @lru_cache(maxsize=None)
        def style(index: int):
            node = groups["cellXfs"][index]
            attrs = dict(node.attrib)
            number = attrs.pop("numFmtId", "0")
            attrs["numberFormat"] = number_formats.get(number, "builtin:" + number)
            references = []
            for attribute, group in (("fontId", "fonts"), ("fillId", "fills"), ("borderId", "borders"), ("xfId", "cellStyleXfs")):
                references.append(canonical(groups[group][int(attrs.pop(attribute, "0"))]))
            return tuple(sorted(attrs.items())), tuple(references), tuple(map(canonical, node))

        book = ET.fromstring(archive.read("xl/workbook.xml"))
        relations = {node.attrib["Id"]: node.attrib["Target"] for node in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))}
        sheet = next(node for node in book.find(S + "sheets") if node.attrib["name"] == "Results")
        target = relations[sheet.attrib[R + "id"]]
        document = ET.fromstring(archive.read(target.lstrip("/") if target.startswith("/") else "xl/" + target))
        cells, rows, columns = {}, {}, []
        for row in document.find(S + "sheetData"):
            attrs = {key: value for key, value in row.attrib.items() if key not in ("r", "spans")}
            if "s" in attrs:
                attrs["s"] = style(int(attrs["s"]))
            rows[row.attrib["r"]] = attrs
            for cell in row:
                cells[cell.attrib["r"]] = style(int(cell.attrib.get("s", "0")))
        for column in document.findall(S + "cols/" + S + "col"):
            attrs = dict(column.attrib)
            if "style" in attrs:
                attrs["style"] = style(int(attrs["style"]))
            columns.append(attrs)
        return dict(cells=cells, rows=rows, columns=columns,
                    sections={tag: tuple(map(canonical, document.findall(S + tag)))
                              for tag in ("mergeCells", "conditionalFormatting", "sheetFormatPr")})


def main() -> None:
    """Проверяет каждый стиль, размер и объединение, независимо от меняющихся RunID и времени."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline", type=Path, required=True)
    parser.add_argument("--candidate", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    a, b = read(args.baseline), read(args.candidate)
    errors = []
    for address in a["cells"].keys() | b["cells"].keys():
        if a["cells"].get(address) != b["cells"].get(address):
            errors.append("Cell style: " + address)
    for row in a["rows"].keys() | b["rows"].keys():
        if a["rows"].get(row) != b["rows"].get(row):
            errors.append("Row formatting: " + row)
    for key in ("columns", "sections"):
        if a[key] != b[key]:
            errors.append(key)
    result = dict(Status="FAIL" if errors else "PASS", CellsCompared=len(a["cells"].keys() | b["cells"].keys()),
                  BaselineSHA256=hashlib.sha256(args.baseline.read_bytes()).hexdigest().upper(),
                  CandidateSHA256=hashlib.sha256(args.candidate.read_bytes()).hexdigest().upper(),
                  Errors=errors)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({key: result[key] for key in ("Status", "CellsCompared", "Errors")}))
    if errors:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
