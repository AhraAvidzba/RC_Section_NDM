"""ДЛЯ ТЕСТОВ: принимает private Ready30 только после полного PASS и native-проверок.

Проверяет сохранность пользовательских ZIP-частей и ячеек независимо от отчёта
подготовки. Не запускает Excel, не пересчитывает и не публикует книгу.
"""
from pathlib import Path
from zipfile import ZipFile
from datetime import datetime, timezone
from copy import deepcopy
import argparse
import hashlib
import json
import re
import posixpath
from lxml import etree as ET
from TestReady30SchemaTools02 import TestCells, TestValue, TestPayload, TestCanonical, TestTag

ROOT = Path('C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM')
PRIVATE = Path(__file__).parent
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def TestJSON(path):
    """Читает нативный либо Python JSON, не меняя исходный файл."""
    return json.loads(path.read_text(encoding='utf-8-sig'))


def TestSHA(path):
    """Вычисляет точный хеш всего файла без нормализации ZIP или строк."""
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def TestExportWithoutDate(path):
    """Исключает только служебную дату экспорта; код и порядок181 компонентов остаются точными."""
    text = path.read_bytes().decode('utf-8-sig')
    assert len(re.findall(r'(?m)^Date: ', text)) == 1
    return re.sub(r'(?m)^Date: [^\r\n]*\r?\n', '', text)


def TestReview(launch_directory):
    """Требует всю приёмку182/122,181native-компонент и завершение точного собственного ReadyExcel."""
    launch = (ROOT / launch_directory).resolve()
    assert launch.is_relative_to(ROOT / 'docs/regression')
    request = TestJSON(launch / 'LaunchRequest.json')
    accept = TestJSON(ROOT / request['EvidenceDirectory'] / 'CompleteCurrentAcceptance.json')
    own = TestJSON(launch / 'AllOwnClosuresReviewed.json')
    assert accept['Status'] == 'PASS' and accept['SourceSHA256'] == request['SourceSHA256']
    assert accept['CoreSuites'] == 182 and accept['FocusedSingleCases'] == 122
    assert own['AllOwnIdentitiesTerminal'] and own['PrimaryLogsReviewed']
    assert own['SourceSHA256'] == request['SourceSHA256'] and own['CoreMacrosReviewed'] == 182
    assert own['FocusedSingleCasesReviewed'] == 122 and own['FocusedRecordCount'] == 26
    prep = TestJSON(PRIVATE / 'Ready30SchemaPreparation02.json')
    native_dir = PRIVATE / 'Ready30SchemaNative02'
    native = TestJSON(native_dir / 'NativeVerification.json')
    quit_proof = TestJSON(native_dir / 'OwnClosure.json')
    equality = TestJSON(PRIVATE / 'Ready30SchemaEquality02.json')
    dropdowns = TestJSON(native_dir / 'DropdownAcceptance.json')
    schema = TestJSON(native_dir / 'SchemaNativeAcceptance.json')
    assert not (native_dir / 'Failure.json').exists()
    assert prep['LaunchDirectory'] == launch_directory and prep['AcceptedMainSHA256'] == request['SourceSHA256']
    assert prep['AcceptedCommit'] == request['SourceCommit'] and not prep['OnlyVBAChanged']
    assert equality['Passed'] and equality['SourceComponents'] == 175 and equality['ExportComponents'] == 181
    assert not equality['Missing'] and not equality['Differences']
    assert quit_proof['ExactIdentityTerminal'] and quit_proof['QuitRequested'] and not quit_proof['ForceKilled']
    assert dropdowns['Status'] == 'PASS_SAVED_READY_NATIVE_DROPDOWNS' and not dropdowns['WorkbookSaved']
    assert schema['MainAcceptedChoices'] == 68 and schema['MainRejectedChoices'] == 85 and schema['ProjectionDashes'] == 4
    assert len(schema['SharedTargets']) == 4
    assert native['Concrete'] == 1232 and native['Rebar'] == 8 and len(native['Components']) == 181
    assert native['LoadIDs'] == [f'PERF_{number:02}' for number in range(1, 31)]
    assert native['HistoricalInputErrIDs'] == ['PERF_14', 'PERF_21', 'PERF_27', 'PERF_28']
    assert native['Settings']['Results.SnapshotStorage'] == 'Excel'
    assert native['Settings']['Output.CombinationMode'] == 'Worst'
    assert native['Settings']['Mesh.CellGeometry'] == 'Adaptive'
    target = Path(prep['Ready'])
    original = Path(prep['SourceBackup'])
    destination = Path(prep['Source'])
    main = ROOT / 'workbook/output/RC_Section_NDM.xlsm'
    assert TestSHA(main) == request['SourceSHA256']
    assert TestSHA(original) == TestSHA(destination) == prep['SourceSHA256']
    assert TestSHA(target) == native['SHA256'] == prep['ReadySHA256']
    assert TestExportWithoutDate(ROOT / 'workbook/output/VBA_All_Code.txt') == TestExportWithoutDate(native_dir / 'VBA_All_Code.txt')
    with ZipFile(original) as old, ZipFile(target) as new:
        assert old.testzip() is None and new.testzip() is None
        old_names, new_names = set(old.namelist()), set(new.namelist())
        assert old_names <= new_names and len(new_names) == len(new.namelist())
        changed = {name for name in old_names if old.read(name) != new.read(name)}
        assert changed == set(prep['ChangedOriginalParts'])
        added = new_names - old_names
        assert added == {image['NewPart'] for image in prep['AddedHelpImages']}
        assert all(name.startswith('xl/media/ndm_help_current_') for name in added)
        # Results и схема вместе со всеми связанными таблицами, стилями по исходным ID и рисунками остаются исходными.
        allowed = {'xl/vbaProject.bin', 'xl/styles.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml', 'xl/workbook.xml', 'xl/worksheets/_rels/sheet3.xml.rels'}
        help_rels = ET.fromstring(old.read('xl/worksheets/_rels/sheet3.xml.rels'))
        assert len(help_rels) == 1
        target_part = help_rels[0].get('Target')
        drawing = target_part.lstrip('/') if target_part.startswith('/') else posixpath.normpath(posixpath.join('xl/worksheets', target_part))
        assert drawing.startswith('xl/') and '..' not in drawing.split('/')
        allowed.add(drawing)
        allowed.add(posixpath.join(posixpath.dirname(drawing), '_rels', posixpath.basename(drawing) + '.rels'))
        assert changed <= allowed, changed - allowed
        for part in ('xl/worksheets/sheet4.xml', 'xl/worksheets/sheet5.xml', 'xl/sharedStrings.xml', 'xl/theme/theme1.xml'):
            assert old.read(part) == new.read(part), part
        old_styles = ET.fromstring(old.read('xl/styles.xml'))
        new_styles = ET.fromstring(new.read('xl/styles.xml'))
        for collection in ('numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs'):
            before = list(old_styles.find(TestTag(collection)))
            after = list(new_styles.find(TestTag(collection)))
            assert len(after) >= len(before)
            assert [TestCanonical(node, True) for node in after[:len(before)]] == [TestCanonical(node, True) for node in before], collection
        old_strings = ET.fromstring(old.read('xl/sharedStrings.xml'))
        new_strings = ET.fromstring(new.read('xl/sharedStrings.xml'))
        config_old = TestCells(ET.fromstring(old.read('xl/worksheets/sheet1.xml')))
        config_new = TestCells(ET.fromstring(new.read('xl/worksheets/sheet1.xml')))
        assert {key: TestPayload(cell) for key, cell in config_old.items()} == {key: TestPayload(cell) for key, cell in config_new.items()}
        geometry_old = TestCells(ET.fromstring(old.read('xl/worksheets/sheet2.xml')))
        geometry_new = TestCells(ET.fromstring(new.read('xl/worksheets/sheet2.xml')))
        machine_rows = sorted(int(address[1:]) for address, cell in geometry_old.items() if address.startswith('H') and TestValue(cell, old_strings) == 'note')
        assert len(machine_rows) == 4
        definitions = TestJSON(ROOT / 'tools/build_workbook/GeometryTables.json')
        allowed_cells = set()
        for row, definition in zip(machine_rows, definitions):
            rows = next(block['Rows'] for block in definition['Blocks'] if block['Key'] == 'Main')
            for number in range(row - 1, row + 2 + len(rows)):
                h, i = 'H' + str(number), 'I' + str(number)
                allowed_cells.update((h, i))
                assert TestPayload(geometry_old[h]) == TestPayload(geometry_new[i])
            for offset, data in enumerate(rows, row + 2):
                assert TestValue(geometry_new['H' + str(offset)], new_strings) == data[7]
        assert set(geometry_old) <= set(geometry_new)
        for address, cell in geometry_old.items():
            if address not in allowed_cells:
                assert TestPayload(cell) == TestPayload(geometry_new[address]), address
        protected_parts = sorted(old_names - changed)
    proof = {'Status': 'PASS_CURRENT_FULL_NATIVE_READY_SCHEMA_DATA_PRESERVATION', 'ObservedUTC': datetime.now(timezone.utc).isoformat(), 'LaunchDirectory': launch_directory, 'AcceptedMainSHA256': request['SourceSHA256'], 'AcceptedCommit': request['SourceCommit'], 'OriginalReadySHA256': prep['SourceSHA256'], 'ReadySHA256': prep['ReadySHA256'], 'Ready': str(target), 'Destination': str(destination), 'NativeComponents': 181, 'MainReadyExportExactExceptDate': True, 'NativeOwnPID': quit_proof['PID'], 'NativeOwnStartTicksUTC': str(quit_proof['StartTicksUTC']), 'NativeOwnTerminal': True, 'ProtectedOriginalParts': protected_parts, 'ProtectedPartCount': len(protected_parts), 'ConfigCellPayloadsExact': True, 'OtherGeometryUserCellsExact': True, 'MainCommentsMovedExact': True, 'OldStylesPrefixExact': True, 'ResultsAndCalculationSheetByteExact': True, 'OriginalThirtyLoadsAndFourInputErrPreserved': True, 'SolverRecalculated': False, 'FullAcceptanceAppliesToMainSourceSHAOnly': True, 'Published': False}
    (PRIVATE / 'Ready30FinalReview03.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({key: proof[key] for key in ('Status', 'AcceptedMainSHA256', 'ReadySHA256', 'ProtectedPartCount', 'Published')}, ensure_ascii=False))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--launch-directory', required=True)
    TestReview(parser.parse_args().launch_directory)
