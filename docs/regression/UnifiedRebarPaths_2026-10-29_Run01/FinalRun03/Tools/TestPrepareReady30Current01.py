"""ДЛЯ ТЕСТОВ: готовит code-only копию настоящей текущей Ready30 после полного PASS.

Не публикует книгу, не запускает Excel и не пересчитывает исторические данные.
Все части пакета, кроме VBA, обязаны остаться побайтно неизменными.
"""
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile
import argparse
import hashlib
import json
import posixpath
import re

ROOT = Path(r"C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM")
PRIVATE = Path(__file__).resolve().parent
READY_SHA = "6A21BE6E842D88317683D0BEB6228587C387014D24A8D42CA721785F664009C0"
OTHER_SHA = "31F188692A01CC421A37CF962598778A5AD1FE6D9AB5D96DD41C29389C0919F7"
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def read_json(path):
    """Читает сохранённое доказательство с допустимым UTF-8 BOM."""
    return json.loads(path.read_text("utf-8-sig"))


def sha(path):
    """Возвращает SHA256 всего пакета, без округления или нормализации."""
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def cell_position(address):
    """Переводит адрес одной ячейки в пару номер-столбца/номер-строки."""
    match = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)", address)
    assert match, address
    letters, row = match.groups()
    column = 0
    for letter in letters:
        column = column * 26 + ord(letter) - 64
    return column, int(row)


def contains(sqref, address):
    """Проверяет попадание ячейки в реальную область проверки данных."""
    column, row = cell_position(address)
    for area in sqref.split():
        ends = area.split(":")
        first, last = cell_position(ends[0]), cell_position(ends[-1])
        if first[0] <= column <= last[0] and first[1] <= row <= last[1]:
            return True
    return False


def contract(path):
    """Снимает совместимость документа и динамических VBA-привязок без Excel.

    Это только предварительный контракт; приёмка выпадающих списков остаётся
    нативной и выполняется после подготовки сохранённого пакета.
    """
    with ZipFile(path) as package:
        assert package.testzip() is None
        strings = []
        if "xl/sharedStrings.xml" in package.namelist():
            strings = ["".join(item.itertext()) for item in ET.fromstring(package.read("xl/sharedStrings.xml"))]
        book = ET.fromstring(package.read("xl/workbook.xml"))
        names = {item.get("name"): item.text for item in book.findall("m:definedNames/m:definedName", NS)}
        rels = {item.get("Id"): item.get("Target") for item in ET.fromstring(package.read("xl/_rels/workbook.xml.rels"))}
        sheets, codenames, grids, validations = [], {}, {}, {}
        for sheet in book.findall("m:sheets/m:sheet", NS):
            name = sheet.get("name")
            sheets.append(name)
            target = rels[sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")]
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            xml = ET.fromstring(package.read(part))
            properties = xml.find("m:sheetPr", NS)
            codenames[name] = properties.get("codeName") if properties is not None else None
            validations[name] = [{"Range": item.get("sqref"), "Type": item.get("type"),
                                  "Formula1": item.findtext("m:formula1", default="", namespaces=NS)}
                                 for item in xml.findall("m:dataValidations/m:dataValidation", NS)]
            grid = {}
            for cell in xml.findall("m:sheetData/m:row/m:c", NS):
                value = cell.find("m:v", NS)
                if cell.get("t") == "inlineStr":
                    text = "".join(cell.find("m:is", NS).itertext())
                elif value is None:
                    text = ""
                elif cell.get("t") == "s":
                    text = strings[int(value.text)]
                else:
                    text = value.text
                grid[cell.get("r")] = text
            grids[name] = grid
        sheet_name, area = names["rngSystemSettings"].split("!", 1)
        sheet_name = sheet_name.strip("'")
        column, first_row = re.match(r"\$?([A-Z]+)\$?(\d+)", area).groups()
        settings = {value: address for address, value in grids[sheet_name].items()
                    if re.fullmatch(column + r"\d+", address) and int(re.search(r"\d+", address)[0]) > int(first_row)
                    and value and not value.startswith("[")}
        value_column = cell_position(column + "1")[0] + 1
        letters = ""
        while value_column:
            value_column, remainder = divmod(value_column - 1, 26)
            letters = chr(65 + remainder) + letters
        fixed = {}
        for key in ("Geometry.Type", "Mesh.CellGeometry", "Mesh.StepMode", "Results.SnapshotStorage", "Output.CombinationMode"):
            address = letters + re.search(r"\d+", settings[key])[0]
            fixed[key] = [{k: v for k, v in item.items() if k != "Range"}
                          for item in validations[sheet_name] if contains(item["Range"], address)]
            assert len(fixed[key]) == 1 and fixed[key][0]["Type"] == "list", key
        critical = ("rngNDMSectionGeometry", "rngNDMSectionProperties", "rngNDMMaterialDiagrams",
                    "rngNDMSectionContours", "rngNDMElementResults", "rngNDMElementGeometry",
                    "rngNDMSectionAnnotations", "rngLoadCombinations", "rngSystemSettings")
        properties = book.find("m:workbookPr", NS)
        return {"Sheets": sheets, "CodeNames": codenames, "WorkbookCodeName": properties.get("codeName"),
                "SettingKeys": sorted(settings), "Anchors": {key: names[key] for key in critical},
                "FixedSettingValidations": fixed}


def main():
    """Допускает подготовку только после точного полного PASS и всех own closures."""
    parser = argparse.ArgumentParser()
    parser.add_argument("--launch-directory", required=True)
    args = parser.parse_args()
    launch = (ROOT / args.launch_directory).resolve()
    assert launch.is_relative_to(ROOT / "docs/regression")
    request = read_json(launch / "LaunchRequest.json")
    complete = read_json(ROOT / request["EvidenceDirectory"] / "CompleteCurrentAcceptance.json")
    own = read_json(launch / "AllOwnClosuresReviewed.json")
    assert complete["Status"] == "PASS" and complete["SourceSHA256"] == request["SourceSHA256"]
    assert complete["CoreSuites"] == 182 and complete["FocusedSingleCases"] == 122
    assert own["AllOwnIdentitiesTerminal"] and own["PrimaryLogsReviewed"]
    assert own["SourceSHA256"] == request["SourceSHA256"]
    assert own["CoreMacrosReviewed"] == 182 and own["FocusedSingleCasesReviewed"] == 122 and own["FocusedRecordCount"] == 26
    main_book = ROOT / "workbook/output/RC_Section_NDM.xlsm"
    source = ROOT / "workbook/output/NDM_Test_1232_30/RC_Section_NDM.xlsm"
    other = source.with_name("RC_Section_NDM_ok.xlsm")
    assert sha(main_book) == request["SourceSHA256"], "Current main differs from the full accepted source"
    assert sha(source) == READY_SHA, "Ready30 has changed; preserve and review it before preparing a replacement"
    assert sha(other) == OTHER_SHA, "The unrelated small example has changed; review its preservation fence"
    main_contract, source_contract = contract(main_book), contract(source)
    assert main_contract == source_contract, "Main/Ready structural contract differs; no output was created"
    output_directory = PRIVATE / "Ready30Current01"
    output_directory.mkdir(exist_ok=False)
    output = output_directory / "RC_Section_NDM.xlsm"
    source_backup = output_directory / "Before_Ready30.xlsm"
    source_backup.write_bytes(source.read_bytes())
    with ZipFile(main_book) as package:
        vba = package.read("xl/vbaProject.bin")
    with ZipFile(source) as old, ZipFile(output, "w") as new:
        names = old.namelist()
        assert len(names) == len(set(names)) == 111
        assert not any("signature" in name.lower() for name in names)
        for item in old.infolist():
            new.writestr(item, vba if item.filename == "xl/vbaProject.bin" else old.read(item.filename))
    preserved = []
    with ZipFile(source) as old, ZipFile(output) as new:
        assert new.testzip() is None and old.namelist() == new.namelist()
        for name in old.namelist():
            if name == "xl/vbaProject.bin":
                assert new.read(name) == vba
            else:
                original = old.read(name)
                assert original == new.read(name), name
                preserved.append({"Part": name, "Bytes": len(original), "SHA256": hashlib.sha256(original).hexdigest().upper()})
    proof = {"Status": "PREPARED_NOT_PUBLISHED_NATIVE_REVIEW_REQUIRED", "LaunchDirectory": args.launch_directory,
             "AcceptedMainSHA256": sha(main_book), "AcceptedCommit": request["SourceCommit"],
             "Source": str(source), "SourceSHA256": sha(source), "SourceBackup": str(source_backup),
             "Ready": str(output), "ReadySHA256": sha(output), "VBAProjectSHA256": hashlib.sha256(vba).hexdigest().upper(),
             "OnlyVBAChanged": True, "PreservedCount": len(preserved), "PreservedParts": preserved,
             "Contract": main_contract, "DataAndLoadsRecalculated": False, "OtherCopyUnchanged": True,
             "FullSourcePassVerified": True, "CurrentFullOwnClosed": True, "Published": False}
    (PRIVATE / "Ready30CurrentPreparation01.json").write_text(json.dumps(proof, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({key: value for key, value in proof.items() if key not in ("PreservedParts", "Contract")}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
