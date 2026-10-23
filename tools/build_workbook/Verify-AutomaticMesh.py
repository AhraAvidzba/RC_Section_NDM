# ==========================================================================
# ДЛЯ ТЕСТОВ: сохранность книги и адресная приёмка Auto/User
# ==========================================================================
"""Проверяет сохранность данных/оформления и адресную приёмку Auto/User."""
import argparse
import hashlib
import json
import re
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestSignature(element):
    """Сравнивает смысл XML без служебных namespace-префиксов."""
    return (E.QName(element).localname, tuple(sorted(element.attrib.items())),
            element.text or '', tuple(TestSignature(child) for child in element))


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestStyles(archive):
    """Разворачивает индексы шрифтов/заливок/рамок, которые Excel сжимает при Save."""
    root = E.fromstring(archive.read('xl/styles.xml'))
    pools = {key: list(root.find('m:' + key, NS)) for key in ('fonts', 'fills', 'borders')}
    formats = {item.get('numFmtId'): item.get('formatCode') for item in root.find('m:numFmts', NS)}
    result = []
    for item in root.find('m:cellXfs', NS):
        attributes = dict(item.attrib)
        for attribute, pool in (('fontId', 'fonts'), ('fillId', 'fills'), ('borderId', 'borders')):
            attributes[attribute] = TestSignature(pools[pool][int(attributes[attribute])])
        attributes['numFmtId'] = formats.get(attributes.get('numFmtId', '0'), attributes.get('numFmtId', '0'))
        result.append((tuple(sorted(attributes.items())), tuple(TestSignature(child) for child in item)))
    return result


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestCells(archive, part):
    """Читает фактические значения, формулы и оформление непустых ячеек."""
    strings = [''.join(item.itertext()) for item in E.fromstring(archive.read('xl/sharedStrings.xml'))]
    cell_styles = TestStyles(archive)
    result = {}
    for cell in E.fromstring(archive.read(part)).findall('.//m:sheetData/m:row/m:c', NS):
        value = cell.find('m:v', NS)
        formula = cell.find('m:f', NS)
        value = value.text if value is not None else ''
        if cell.get('t') == 's':
            value = strings[int(value)]
        elif cell.get('t') == 'inlineStr':
            value = ''.join(cell.find('m:is', NS).itertext())
        elif value and cell.get('t') not in ('str', 'e'):
            value = float(value)
        if value != '' or formula is not None:
            result[cell.get('r')] = (formula.text if formula is not None else '', value,
                                    cell_styles[int(cell.get('s', '0'))])
    return result


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestMoved(address, config=False):
    """Отражает ровно согласованные вставки: A:E Config и девять строк Справки."""
    def replace(match):
        column, row = match[1], int(match[2])
        if config:
            row += int(column.replace('$', '') in ('A', 'B', 'C', 'D', 'E') and row >= 16)
        else:
            row += 9 if row >= 799 else 0
        return column + str(row)
    return re.sub(r'(\$?[A-Z]+\$?)(\d+)', replace, address)


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestMain():
    """Требует обе адресные группы ON/OFF и нативную проверку той же книги."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory', default='docs/regression/AutomaticMesh_2026-10-23_Run03')
    args = parser.parse_args()
    directory = Path(args.directory)
    target = directory / 'RC_Section_NDM.xlsm'
    sha = hashlib.sha256(target.read_bytes()).hexdigest().upper()
    native = json.loads((directory / 'WorkbookVerification.json').read_text(encoding='utf-8-sig'))
    assert native['Status'] == 'PASS' and native['WorkbookSHA256'] == sha
    checked = {}
    with ZipFile(directory / 'Before.xlsm') as old, ZipFile(target) as new:
        for number in (2, 4, 5):
            part = f'xl/worksheets/sheet{number}.xml'
            original, current = TestCells(old, part), TestCells(new, part)
            assert original == current, f'Unrelated values/formulas/styles changed: {part}'
            checked[part] = len(original)
        original, current = TestCells(old, 'xl/worksheets/sheet1.xml'), TestCells(new, 'xl/worksheets/sheet1.xml')
        allowed = {'D17', 'D18', 'D19', 'B19'}
        picker_cells = {'BS' + str(row) for row in range(2, 32)}
        for address, value in original.items():
            destination = TestMoved(address, True)
            assert destination in current
            if destination in picker_cells:
                assert value[2] == current[destination][2]
                assert current[destination][0] == 'IF(ROW()-1>ROWS(rngLoadCombinations)-1,"",INDEX(rngLoadCombinations,ROW(),1))'
            elif destination in allowed:
                assert value[0] == current[destination][0] and value[2] == current[destination][2]
            else:
                assert value == current[destination], f'Config changed: {address} -> {destination}'
        assert set(current) - {TestMoved(address, True) for address in original} == {'A16', 'B16', 'C16', 'D16', 'E16'}
        checked['ConfigPreservedCells'] = len(original)
        checked['LoadCasePickerRowsVerifiedInExcel'] = native['LoadCasePickerRows']
        original, current = TestCells(old, 'xl/worksheets/sheet3.xml'), TestCells(new, 'xl/worksheets/sheet3.xml')
        allowed_help = set()
        for key in ('Mesh.StepX', 'Mesh.StepY', 'Mesh.BoundarySubdivisions'):
            address = next(address for address, value in original.items() if value[1] == key)
            row = int(re.search(r'\d+', address)[0]) + 1
            while original.get('A' + str(row), ('', ''))[1] != '':
                allowed_help.add(TestMoved('B' + str(row)))
                row += 1
        for address, value in original.items():
            destination = TestMoved(address)
            assert destination in current
            if destination in allowed_help:
                assert value[0] == current[destination][0] and value[2] == current[destination][2]
            else:
                assert value == current[destination], f'Help changed: {address}'
        assert set(current) - {TestMoved(address) for address in original} == {'A799'} | {'A' + str(row) for row in range(800, 807)} | {'B' + str(row) for row in range(800, 807)}
        checked['HelpPreservedCells'] = len(original)
        for number in (1, 2, 3, 4, 5):
            part = f'xl/worksheets/sheet{number}.xml'
            previous = E.fromstring(old.read(part))
            current = E.fromstring(new.read(part))
            links = {item.get('ref'): item.get('location') for item in current.findall('m:hyperlinks/m:hyperlink', NS)}
            for item in previous.findall('m:hyperlinks/m:hyperlink', NS):
                ref = item.get('ref')
                if number == 1:
                    ref = TestMoved(ref, True)
                elif number == 3:
                    ref = TestMoved(ref)
                destination = item.get('location')
                if destination and 'Справка' in destination:
                    prefix, address = destination.rsplit('!', 1)
                    destination = prefix + '!' + TestMoved(address)
                assert links.get(ref, '').replace("'", '') == (destination or '').replace("'", ''), f'Link changed: {part} {ref}'
        media = [name for name in old.namelist() if name.startswith('xl/media/')]
        assert all(old.read(name) == new.read(name) for name in media)
        checked['UnchangedMediaParts'] = len(media)
        diagram_parts = [name for name in old.namelist()
                         if (name.startswith('xl/drawings/') or name.startswith('xl/charts/'))
                         and name != 'xl/drawings/drawing2.xml']
        assert all(old.read(name) == new.read(name) for name in diagram_parts)
        checked['UnchangedDrawingAndChartParts'] = len(diagram_parts)
        # -------------------- ДЛЯ ТЕСТОВ --------------------
        def TestNames(archive):
            return {(item.get('name'), item.get('localSheetId')): (item.text or '').replace("'", '').lstrip('=')
                    for item in E.fromstring(archive.read('xl/workbook.xml')).findall('m:definedNames/m:definedName', NS)}
        previous, current = TestNames(old), TestNames(new)
        assert current.pop(('rngSystemSettings', None)) == 'Config!$A$3:$E$140'
        previous.pop(('rngSystemSettings', None))
        assert current == previous, 'Unrelated names changed'
    suites = {}
    for mode in ('No', 'Yes'):
        report = (directory / ('Final_' + mode + '.txt')).read_text(encoding='utf-8-sig')
        assert 'SOURCE_SHA256: ' + sha in report and 'WATCHDOG_COMPLETED: exit=0' in report
        assert not re.search(r'^FAIL:', report, re.M)
        totals = re.findall(r'TOTAL_(AUTO_MESH|CONTOUR_MESH|GENERAL_GEOMETRY_CONFIG|WORKFLOW_SNAPSHOT): passed=(\d+); failed=(\d+)', report)
        assert len(totals) == 4 and len({item[0] for item in totals}) == 4
        assert all(int(passed) > 0 and int(failed) == 0 for _, passed, failed in totals)
        suites[mode] = {'passed': sum(int(item[1]) for item in totals), 'failed': 0, 'groups': 4}
    receipt = {'Status': 'PASS', 'WorkbookSHA256': sha, 'Preservation': checked, 'Directed': suites,
               'FullOnOffExecuted': False, 'NativeWorkbookVerification': native}
    (directory / 'Acceptance.json').write_text(json.dumps(receipt, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'Status': 'PASS', 'WorkbookSHA256': sha, 'Directed': suites}, ensure_ascii=False))


if __name__ == '__main__':
    TestMain()
