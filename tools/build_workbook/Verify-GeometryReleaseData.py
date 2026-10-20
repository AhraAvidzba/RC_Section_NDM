"""Проверяет исходные пользовательские данные с точной картой миграции Config."""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("recipe_preservation", Path(__file__).with_name("Verify-RecipeConfigPreservation.py"))
preserve = importlib.util.module_from_spec(spec); spec.loader.exec_module(preserve)


def bounds(reference):
    match = re.search(r"!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)$", reference)
    if not match: raise ValueError("Expected a rectangular named range: " + reference)
    return int(match[2]), preserve.column_number(match[1]), int(match[4]), preserve.column_number(match[3])


def verify(before_path, after_path):
    before, after = preserve.read_book(before_path), preserve.read_book(after_path)
    expected = {sheet: dict(values) for sheet, values in before["sheets"].items()}
    config = expected["Config"]
    allowed = set()
    def change(row, column, value):
        config[(row, column)] = value; allowed.add((row, column))
    for name, count in (("rngPointedSectionGeometry", 4), ("rngHollowRectangleGeometry", 8), ("rngRectanglesGeometry", 0)):
        top, left, _, _ = bounds(before["names"][name])
        data = preserve.table(before, name)
        main = next(index for index, row in enumerate(data) if len(row) > 1 and row[1][1] == "as")
        for column, text in enumerate(("StartSetback", "EndSetback"), 4):
            change(top+main, left+column, (None, text))
        if count:
            change(top+main, left+6, (None, "Ед.")); change(top+main, left+7, (None, "Комментарий"))
            for row in range(main+1, main+count+1):
                for column in (4, 5): change(top+row, left+column, (None, "0"))
                for column in (6, 7): change(top+row, left+column, data[row][column-2])
    name = "rngCircleGeometry"
    top, left, _, _ = bounds(before["names"][name]); data = preserve.table(before, name)
    row = next(index for index, values in enumerate(data) if values[0][1] == "Rebar.Count")
    explanation = "ClosedUniform: StartSetback/EndSetback неприменимы к полному кольцу. Для отступов задайте открытый участок дуги в рецепте."
    change(top+row, left+3, (None, data[row][3][1] + " " + explanation))
    types = re.findall(r'(?m)^\s*RegisterRecipe "([^\"]+)"', (ROOT/"src/Excel/CSectionTypeRegistry.cls").read_text(encoding="utf-8-sig"))
    top, left, bottom, right = bounds(after["names"]["rngSectionRecipeTypes"])
    if right != left or bottom-top+1 != len(types): raise ValueError("Type-list bounds mismatch")
    for index, type_id in enumerate(types): change(top+index, left, (None, type_id))
    errors, changes, checks = [], [], 0
    added = sorted(set(after["names"]) - set(before["names"]))
    removed = sorted(set(before["names"]) - set(after["names"]))
    if added != ["rngSectionRecipeGuide", "rngSectionRecipeTypes"] or removed:
        errors.append("Unexpected named range delta")
    for name, reference in before["names"].items():
        if after["names"].get(name) != reference: errors.append("Old name changed: " + name)
    for sheet, cells in expected.items():
        if sheet == "Справка": continue
        for position in set(cells) | set(after["sheets"][sheet]):
            actual = after["sheets"][sheet].get(position, (None, None))
            value = cells.get(position, (None, None))
            if value != (None, None) or actual != (None, None):
                checks += 1
                if actual != value: errors.append(f"Unexpected user value/formula change: {sheet}!{position}: {value!r} -> {actual!r}")
            original = before["sheets"][sheet].get(position, (None, None))
            if actual != original:
                changes.append({"sheet": sheet, "row": position[0], "column": position[1], "before": original, "after": actual})
    help_cells = after["sheets"]["Справка"]
    guide_top, _, guide_bottom, _ = bounds(after["names"]["rngSectionRecipeGuide"])
    for position, value in before["sheets"]["Справка"].items():
        if help_cells.get(position, (None, None)) != value: errors.append("Old Help text/formula changed: " + str(position))
    for position, value in help_cells.items():
        if position[0] > guide_bottom and value != (None, None): errors.append("Content follows recipe guide")
    return {"schema": 1, "status": "USER_DATA_PRESERVED" if not errors else "FAIL", "checkedNonEmptyCells": checks,
            "beforeSHA256": hashlib.sha256(before_path.read_bytes()).hexdigest(), "afterSHA256": hashlib.sha256(after_path.read_bytes()).hexdigest(),
            "guideStart": guide_top, "guideEnd": guide_bottom, "addedNames": added, "removedNames": removed,
            "printAreaPreserved": before["names"].get("_xlnm.Print_Area") == after["names"].get("_xlnm.Print_Area"),
            "allowedConfigCells": len(allowed), "actualChangesOutsideHelp": changes,
            "loadCombinations": preserve.table(before, "rngLoadCombinations") == preserve.table(after, "rngLoadCombinations"), "errors": errors}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("before", type=Path); parser.add_argument("after", type=Path)
    parser.add_argument("--report", required=True, type=Path)
    args = parser.parse_args(); result = verify(args.before, args.after)
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({key: result[key] for key in ("status", "checkedNonEmptyCells", "allowedConfigCells", "addedNames", "printAreaPreserved", "loadCombinations", "errors")}, ensure_ascii=False))
    if result["errors"]: raise SystemExit(1)
