"""Точечный выпуск автоматической схемы без изменения пользовательских данных."""
from copy import copy
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
import json
import re

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
SOURCE = ROOT / "workbook/output/RC_Section_NDM.xlsm"
DONOR = ROOT / "docs/regression/FormationAutoPreview_2026-10-21/Candidate01/RC_Section_NDM.xlsm"
TARGET = HERE / "Accepted.xlsm"
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def digest(data):
    return sha256(data).hexdigest().upper()


assert not TARGET.exists()
expected = "D894BBB470C1CEAA07B5CD5203A46C8462BEB4512E69ABF605184172C536C031"
assert digest(SOURCE.read_bytes()) == expected, "Пользовательская книга изменилась."
with ZipFile(SOURCE) as old, ZipFile(DONOR) as donor:
    original = {name: old.read(name) for name in old.namelist()}
    updated = original.copy()
    updated["xl/vbaProject.bin"] = donor.read("xl/vbaProject.bin")
    strings = ["".join(item.itertext()) for item in ET.fromstring(original["xl/sharedStrings.xml"])]
    xml = original["xl/worksheets/sheet3.xml"]
    new_text = next(text for text in re.findall(r'"([^"\r\n]+)"', (ROOT / "tools/build_workbook/SettingsCatalog.ps1").read_text(encoding="utf-8-sig"))
                    if text.startswith("Для любого источника сначала нажмите"))
    cells = []
    for cell in ET.fromstring(xml).findall(".//m:sheetData/m:row/m:c", NS):
        text = "".join(cell.find("m:is", NS).itertext()) if cell.get("t") == "inlineStr" else ""
        if cell.get("t") == "s":
            text = strings[int(cell.find("m:v", NS).text)]
        if text.startswith("Для любого источника сначала нажмите"):
            cells.append((cell.get("r"), text))
    assert len(cells) == 1
    address, old_text = cells[0]
    pattern = rb'<c\b[^>]*\br="' + address.encode() + rb'"[^>]*>.*?</c>'
    match = re.search(pattern, xml, re.S)
    opening = re.search(rb"<c\b[^>]*>", match.group()).group()
    opening = re.sub(rb'\bt="[^"]*"', b't="inlineStr"', opening)
    replacement = opening + b'<is><t xml:space="preserve">' + escape(new_text).encode() + b"</t></is></c>"
    updated["xl/worksheets/sheet3.xml"] = xml[:match.start()] + replacement + xml[match.end():]
    ET.fromstring(updated["xl/worksheets/sheet3.xml"])
    changed = [name for name in original if original[name] != updated[name]]
    assert set(changed) == {"xl/vbaProject.bin", "xl/worksheets/sheet3.xml"}
    with ZipFile(TARGET, "w") as out:
        for item in old.infolist():
            out.writestr(copy(item), updated[item.filename])
result = {"Passed": True, "AllInputAndResultCellsPreserved": True, "OriginalSHA256": expected,
          "AcceptedSHA256": digest(TARGET.read_bytes()), "ChangedParts": changed,
          "UserExecutionReportSHA256": digest((ROOT / "workbook/output/RC_Section_NDM_execution_report.txt").read_bytes()),
          "HelpCell": address, "HelpBefore": old_text, "HelpAfter": new_text}
(HERE / "Package.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
print(json.dumps({key: result[key] for key in ["Passed", "AcceptedSHA256", "ChangedParts"]}))
