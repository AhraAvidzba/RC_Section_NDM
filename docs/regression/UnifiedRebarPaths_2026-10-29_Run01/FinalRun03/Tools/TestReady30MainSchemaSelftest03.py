"""ДЛЯ ТЕСТОВ: проверяет защиту пользовательских полей при будущей миграции Main.

Синтетические XML не сохраняются как книга. Реальные книги читаются только
для пассивной проверки совместимости; guarded подготовка и Excel не вызываются.
"""
from pathlib import Path
from copy import deepcopy
from datetime import datetime, timezone
from zipfile import ZipFile
import hashlib
import json
from lxml import etree as ET
from TestReady30SchemaTools02 import TestTag, TestCells, TestPatchMain, TestValue, TestPayload
from TestPrepareReady30Current01 import contract

PRIVATE = Path(__file__).parent
ROOT = Path('C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM')
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def TestSheet(mode='IncludeBoth'):
    """Создаёт независимые старый пользовательский и новый шаблонный Main с различающимися числовыми полями."""
    old = ET.Element(TestTag('worksheet'))
    new = ET.Element(TestTag('worksheet'))
    for sheet, current in ((old, False), (new, True)):
        data = ET.SubElement(sheet, TestTag('sheetData'))
        values = {
            3: {'A': 'Ключ', 'H': 'Расстановка' if current else 'Комментарий'},
            4: {'A': 'key', 'H': 'endMode' if current else 'note'},
            5: {'A': 'Units', 'H': '—' if current else ''},
            6: {'A': 'Side', 'C': '40' if current else '63', 'F': '40' if current else '80', 'H': mode if current else '  Моя пометка Ø32; s₀\nвторая строка  '},
            7: {'A': 'OpeningTop', 'H': '—' if current else 'Комментарий отверстия'},
        }
        for row_number, fields in values.items():
            row = ET.SubElement(data, TestTag('row'), r=str(row_number))
            for column, value in fields.items():
                cell = ET.SubElement(row, TestTag('c'), r=column + str(row_number), s='1' if current else '4', t='inlineStr')
                text = ET.SubElement(ET.SubElement(cell, TestTag('is')), TestTag('t'))
                text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                text.text = value
        checks = ET.SubElement(sheet, TestTag('dataValidations'), count='1')
        rule = ET.SubElement(checks, TestTag('dataValidation'), sqref='H6' if current else 'F6', type='list' if current else 'decimal')
        ET.SubElement(rule, TestTag('formula1')).text = '"IncludeBoth,ExcludeBoth,IncludeStart,IncludeEnd"' if current else '0'
    definitions = [{'Type': 'Synthetic', 'Blocks': [{'Key': 'Main', 'Rows': [['Side'], ['OpeningTop']]}]}]
    return old, new, definitions


def TestReject(change):
    """Требует отказа до публикации при пользовательском заполнении резервного поля или неоднозначном списке."""
    old, new, definitions = TestSheet()
    change(old, new)
    try:
        TestPatchMain(old, new, [], [], {1: 7}, definitions)
    except AssertionError:
        return
    raise AssertionError('Unsafe synthetic Main change was accepted')


def TestReservedCell(old, formula=False):
    """Добавляет пользовательский текст либо формулу в предполагаемый резервный столбец I."""
    row = next(row for row in old.find(TestTag('sheetData')) if row.get('r') == '6')
    cell = ET.SubElement(row, TestTag('c'), r='I6')
    if formula:
        ET.SubElement(cell, TestTag('f')).text = '3+4'
    else:
        cell.set('t', 'inlineStr')
        ET.SubElement(ET.SubElement(cell, TestTag('is')), TestTag('t')).text = 'Пользовательские данные'


def TestSyntheticCases():
    """Проверяет четыре режима, точный перенос комментария/формулы и запрет скрытой замены пользовательских чисел."""
    checks = []
    for mode in ('IncludeBoth', 'ExcludeBoth', 'IncludeStart', 'IncludeEnd'):
        old, new, definitions = TestSheet(mode)
        cells = TestCells(old)
        comment = TestValue(cells['H6'], [])
        outside = {key: TestPayload(cells[key]) for key in ('A6', 'A7', 'C6', 'F6')}
        original_validation = ET.tostring(old.find(TestTag('dataValidations'))[0])
        proof = TestPatchMain(old, new, [], [], {1: 7}, definitions)
        cells = TestCells(old)
        assert TestValue(cells['H6'], []) == mode and cells['H6'].get('s') == '7'
        assert TestValue(cells['H7'], []) == '—'
        assert TestValue(cells['I6'], []) == comment
        assert TestValue(cells['F6'], []) == '80' and TestValue(cells['C6'], []) == '63'
        assert {key: TestPayload(cells[key]) for key in outside} == outside
        assert ET.tostring(old.find(TestTag('dataValidations'))[0]) == original_validation
        assert proof['FixedListModeCells'] == 1 and not proof['TableRowsAndColumnsInserted']
        checks.append(mode + ': exact user values/comment/styles/validation retained')
    old, new, definitions = TestSheet()
    original = TestCells(old)['H6']
    original.clear()
    original.set('r', 'H6')
    original.set('s', '4')
    ET.SubElement(original, TestTag('f')).text = 'SUM(1,2)'
    ET.SubElement(original, TestTag('v')).text = '3'
    payload = TestPayload(original)
    TestPatchMain(old, new, [], [], {1: 7}, definitions)
    assert TestPayload(TestCells(old)['I6']) == payload
    checks.append('User comment formula and cached value retained')
    TestReject(lambda old, new: TestReservedCell(old))
    TestReject(lambda old, new: TestReservedCell(old, True))
    TestReject(lambda old, new: old.find(TestTag('dataValidations'))[0].set('sqref', 'H6'))
    TestReject(lambda old, new: setattr(new.find(TestTag('dataValidations'))[0].find(TestTag('formula1')), 'text', '"IncludeBoth;ExcludeBoth;IncludeStart;IncludeEnd"'))
    checks += ['Reserved user text rejected', 'Reserved user formula rejected', 'Existing user note validation rejected', 'Wrong list separator rejected without normalization']
    return checks


def TestPassiveCompatibility():
    """Читает только XML двух исходных пакетов; книги и процессы не создаёт."""
    main = ROOT / 'workbook/output/RC_Section_NDM.xlsm'
    ready = ROOT / 'workbook/output/NDM_Test_1232_30/RC_Section_NDM.xlsm'
    assert contract(main) == contract(ready)
    result = {'ContractsEqual': True, 'ActualPreparationRun': False, 'ExcelStarted': False, 'BooksCreated': False}
    for key, path in (('Main', main), ('Ready', ready)):
        with ZipFile(path) as package:
            xml = ET.fromstring(package.read('xl/worksheets/_rels/sheet3.xml.rels'))
            targets = [node.get('Target') for node in xml]
            result[key] = {'SHA256': hashlib.sha256(path.read_bytes()).hexdigest().upper(), 'Parts': len(package.namelist()), 'HelpDrawingTargets': targets}
    assert result['Main']['HelpDrawingTargets'] == result['Ready']['HelpDrawingTargets']
    assert result['Main']['SHA256'] == '8BA463BF1AEC97834F81C0D01ADD5707D6D0A21266B6FAE61AA2BCB6608F84E2'
    assert result['Ready']['SHA256'] == '6A21BE6E842D88317683D0BEB6228587C387014D24A8D42CA721785F664009C0'
    return result


if __name__ == '__main__':
    checks = TestSyntheticCases()
    passive = TestPassiveCompatibility()
    proof = {'Status': 'SYNTHETIC_MAIN_DATA_PROTECTION_AND_PASSIVE_COMPATIBILITY_PASS', 'ObservedUTC': datetime.now(timezone.utc).isoformat(), 'SyntheticChecks': checks, 'PassiveCompatibility': passive, 'EndOnlyPreparationOrNativeRun': False, 'FullAcceptanceClaim': False}
    (PRIVATE / 'Ready30MainSchemaSelftest03.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'Status': proof['Status'], 'SyntheticChecks': len(checks), 'ContractsEqual': True, 'ActualPreparationRun': False, 'ExcelStarted': False, 'FullAcceptanceClaim': False}))
