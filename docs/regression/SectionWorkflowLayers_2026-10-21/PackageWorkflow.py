"""Сохраняет пользовательскую книгу, меняя только код и утверждённый интерфейс."""
from hashlib import sha256
from copy import copy
from pathlib import Path
import json
import math
import re
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[3]
HERE = Path(__file__).resolve().parent
ORIGINAL = ROOT / "workbook/output/RC_Section_NDM.xlsm"
DONOR = HERE / "Release20/RC_Section_NDM.xlsm"
TARGET = HERE / "Accepted21.xlsm"
EXPECTED = "BD614F467E673FFAF0B9E02D17C0FF12C68D8668BF740FE7478D74A6DD957434"
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "x": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def digest(value):
    return sha256(value).hexdigest().upper()


def paragraphs():
    source = (ROOT / "tools/build_workbook/SettingsCatalog.ps1").read_text(encoding="utf-8-sig")
    return re.findall(r'"([^"\r\n]+)"', source)


def one(lines, prefix):
    matches = [text for text in lines if text.startswith(prefix)]
    assert len(matches) == 1, (prefix, len(matches))
    return matches[0]


def patch_buttons(original, donor):
    before = original
    targets = ["btnImportGeometryFromAutoCAD", "btnRunSectionCalculation",
               "btnUpdateSectionPlot", "btnExportStressToAutoCAD", "btnClearAutoCADDrawing"]
    pattern = rb"<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>"
    old = {name: next(block for block in re.findall(pattern, original, re.S)
                     if ('name="' + name + '"').encode() in block) for name in targets}
    new = {name: next(block for block in re.findall(pattern, donor, re.S)
                     if ('name="' + name + '"').encode() in block) for name in targets}
    for name in targets:
        block = old[name]
        # Оставляем прежнюю фигуру; изменяем положение, обработчик и надпись.
        for tag in ("xdr:from", "xdr:to", "a:xfrm", "xdr:txBody"):
            part = rb"<" + tag.encode() + rb"\b[^>]*>.*?</" + tag.encode() + rb">"
            match = re.search(part, new[name], re.S)
            assert match, (name, tag)
            block, count = re.subn(part, lambda _: match.group(), block, count=1, flags=re.S)
            assert count == 1
        macro = re.search(rb'<xdr:sp\b[^>]*\bmacro="([^"]*)"', new[name]).group(1)
        block, count = re.subn(rb'(<xdr:sp\b[^>]*\bmacro=")[^"]*(")',
                               lambda match: match[1] + macro + match[2], block, count=1)
        assert count == 1
        original = original.replace(old[name], block, 1)
    # Все остальные фигуры, включая диаграмму, сохраняются побайтно.
    before_other = [b for b in re.findall(pattern, before, re.S)
                    if not any(('name="' + n + '"').encode() in b for n in targets)]
    assert all(b in original for b in before_other)
    ET.fromstring(original)
    return original


def patch_help(xml, strings):
    lines = paragraphs()
    mapping = {
        "AutoCAD: заранее нажмите": one(lines, "Для любого источника сначала нажмите"),
        "При последующем нажатии 'Выполнить расчет'": one(lines, "Кнопка «Выполнить расчет» читает только"),
        "Смена OUTPUT не меняет расчетное НДС.": one(lines, "Смена OUTPUT не меняет расчетное НДС."),
        "Сочетания задаются в таблице Config": one(lines, "Сначала нажмите «Сформировать сечение»:") + " " + one(lines, "CombinationID - идентификатор сочетания;"),
        "2. Нажмите «Импортировать геометрию из AutoCAD».": one(lines, "2. Нажмите «Сформировать сечение». Проверьте"),
        "3. Задайте сочетания и их расчетные профили": one(lines, "3. Задайте сочетания и их расчетные профили"),
    }
    edited = []
    changes = {}
    root = ET.fromstring(xml)
    for cell in root.findall(".//m:sheetData/m:row/m:c", NS):
        if cell.get("t") != "s":
            continue
        value = cell.find("m:v", NS)
        if value is None:
            continue
        old = strings[int(value.text)]
        for prefix, text in mapping.items():
            if old.startswith(prefix):
                address = cell.get("r")
                assert prefix not in changes, prefix
                changes[prefix] = address
                pattern = rb'<c\b[^>]*\br="' + address.encode() + rb'"[^>]*>.*?</c>'
                match = re.search(pattern, xml, re.S)
                assert match, address
                opening = re.search(rb"<c\b[^>]*>", match.group()).group()
                opening = re.sub(rb'\bt="s"', b't="inlineStr"', opening)
                replacement = opening + b'<is><t xml:space="preserve">' + escape(text).encode() + b"</t></is></c>"
                xml = xml[:match.start()] + replacement + xml[match.end():]
                # Строки таблицы настроек имеют перенос и растут только при подготовке книги.
                if address.startswith("B"):
                    row = re.search(r"\d+", address).group()
                    height = max(25.5, math.ceil(len(text) / 125) * 12.75)
                    pattern = rb'<row\b(?=[^>]*\br="' + row.encode() + rb'")[^>]*>'
                    match = re.search(pattern, xml)
                    opening = match.group()
                    if b' ht="' in opening:
                        opening = re.sub(rb'\bht="[^"]*"', ('ht="' + str(height) + '"').encode(), opening)
                    else:
                        opening = opening[:-1] + (' ht="' + str(height) + '" customHeight="1">').encode()
                    xml = xml[:match.start()] + opening + xml[match.end():]
                edited.append({"Cell": address, "Before": old, "After": text})
                break
    assert len(changes) == len(mapping), changes
    ET.fromstring(xml)
    return xml, edited


def main():
    assert not TARGET.exists(), TARGET
    assert digest(ORIGINAL.read_bytes()) == EXPECTED, "Пользовательская книга изменилась."
    with ZipFile(ORIGINAL) as old, ZipFile(DONOR) as donor:
        originals = {name: old.read(name) for name in old.namelist()}
        parts = originals.copy()
        parts["xl/vbaProject.bin"] = donor.read("xl/vbaProject.bin")
        workbook = parts["xl/workbook.xml"]
        assert b"rngNDMSectionModelMetadata" not in workbook
        addition = b'<definedName name="rngNDMSectionModelMetadata">Results!$CO$156</definedName>'
        assert workbook.count(b"</definedNames>") == 1
        parts["xl/workbook.xml"] = workbook.replace(b"</definedNames>", addition + b"</definedNames>")
        parts["xl/drawings/drawing3.xml"] = patch_buttons(parts["xl/drawings/drawing3.xml"], donor.read("xl/drawings/drawing3.xml"))
        strings = ["".join(item.itertext()) for item in ET.fromstring(parts["xl/sharedStrings.xml"])]
        parts["xl/worksheets/sheet3.xml"], help_changes = patch_help(parts["xl/worksheets/sheet3.xml"], strings)
        with ZipFile(TARGET, "w") as out:
            for item in old.infolist():
                out.writestr(copy(item), parts[item.filename])
        changed = [name for name in originals if originals[name] != parts[name]]
        assert set(changed) == {"xl/vbaProject.bin", "xl/workbook.xml", "xl/drawings/drawing3.xml", "xl/worksheets/sheet3.xml"}, changed
        unchanged = {name: digest(parts[name]) for name in old.namelist() if name not in changed}
        for name in ["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml", "xl/worksheets/sheet4.xml", "xl/worksheets/sheet5.xml", "xl/styles.xml"]:
            assert originals[name] == parts[name], name
    report = {"Passed": True, "OriginalSHA256": EXPECTED, "DonorSHA256": digest(DONOR.read_bytes()),
              "AcceptedSHA256": digest(TARGET.read_bytes()), "ChangedParts": changed,
              "AllInputAndResultCellsPreserved": True, "UnchangedParts": unchanged, "HelpChanges": help_changes}
    (HERE / "Package21.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({k: report[k] for k in ["Passed", "AcceptedSHA256", "ChangedParts", "AllInputAndResultCellsPreserved"]}))


if __name__ == "__main__":
    main()
