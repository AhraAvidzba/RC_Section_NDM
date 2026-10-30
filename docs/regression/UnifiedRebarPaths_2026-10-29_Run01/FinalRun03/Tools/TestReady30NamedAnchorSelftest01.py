"""ДЛЯ ТЕСТОВ: проверяет выбор Main по именованной таблице при одинаковом префиксе ключей."""
from copy import deepcopy
from pathlib import Path
from datetime import datetime, timezone
import json
import re
from lxml import etree as ET
from TestReady30MainSchemaSelftest03 import TestSheet, TestSyntheticCases
from TestReady30SchemaTools02 import TestTag, TestCells, TestPatchMain, TestValue
from TestPrepareReady30Schema02 import TestGeometryBounds


def TestAmbiguousSheets():
    """ДЛЯ ТЕСТОВ: создаёт два Main, где ключи первой таблицы являются началом второй."""
    old, new, definitions = TestSheet()
    first = deepcopy(definitions[0]); first['RangeName'] = 'rngFirst'
    second = deepcopy(first); second['Type'] = 'SyntheticLong'; second['RangeName'] = 'rngSecond'
    second['Blocks'][0]['Rows'].append(['OpeningBottom'])
    for sheet in (old, new):
        data = sheet.find(TestTag('sheetData'))
        for original in list(data):
            row = deepcopy(original); row.set('r', str(int(row.get('r')) + 16))
            for cell in row:
                cell.set('r', re.sub(r'\d+$', lambda match: str(int(match[0]) + 16), cell.get('r')))
            data.append(row)
        row = ET.SubElement(data, TestTag('row'), r='24')
        for column, value in (('A', 'OpeningBottom'), ('H', '—' if sheet is new else 'Мой комментарий')):
            cell = ET.SubElement(row, TestTag('c'), r=column + '24', t='inlineStr', s='1' if sheet is new else '4')
            ET.SubElement(ET.SubElement(cell, TestTag('is')), TestTag('t')).text = value
        checks = sheet.find(TestTag('dataValidations'))
        clone = deepcopy(checks[0]); clone.set('sqref', 'H22' if sheet is new else 'F22')
        checks.append(clone); checks.set('count', str(len(checks)))
    return old, new, [first, second]


def TestNamedAnchors():
    """ДЛЯ ТЕСТОВ: неоднозначность отклоняется, точные разные диапазоны выбирают каждый блок один раз."""
    checks = TestSyntheticCases()
    old, new, definitions = TestAmbiguousSheets()
    try:
        TestPatchMain(old, new, [], [], {1: 7}, definitions)
    except AssertionError:
        checks.append('Prefix-only ambiguous Main lookup rejected')
    else:
        raise AssertionError('Ambiguous two-table input accepted')
    old, new, definitions = TestAmbiguousSheets()
    bounds = {'rngFirst': [3, 7], 'rngSecond': [19, 24]}
    proof = TestPatchMain(old, new, [], [], {1: 7}, definitions, bounds)
    assert proof['FixedListModeCells'] == 2 and proof['NamedSectionBounds'] == bounds
    assert TestValue(TestCells(old)['I24'], []) == 'Мой комментарий'
    assert TestValue(TestCells(old)['H24'], []) == '—'
    checks.append('Shared key prefix resolved by exact named-table bounds with comments preserved')
    old, new, definitions = TestAmbiguousSheets()
    try:
        TestPatchMain(old, new, [], [], {1: 7}, definitions, {'rngFirst': [3, 6], 'rngSecond': [19, 24]})
    except AssertionError:
        checks.append('Truncated named-table boundary rejected')
    else:
        raise AssertionError('Main rows outside named table accepted')
    book = ET.Element(TestTag('workbook')); names = ET.SubElement(book, TestTag('definedNames'))
    for name, (begin, end) in bounds.items():
        ET.SubElement(names, TestTag('definedName'), name=name).text = "'Сечения'!$A$%d:$K$%d" % (begin, end)
    assert TestGeometryBounds(book, definitions) == bounds
    names[0].text = "'Results'!$A$3:$K$7"
    try:
        TestGeometryBounds(book, definitions)
    except AssertionError:
        checks.append('Foreign named sheet rejected')
    else:
        raise AssertionError('Foreign named sheet accepted')
    proof = {'Status': 'NAMED_MAIN_ANCHOR_AND_DATA_PROTECTION_SELFTEST_PASS',
             'ObservedUTC': datetime.now(timezone.utc).isoformat(), 'Checks': checks,
             'RealWorkbookChanged': False, 'ExcelStarted': False, 'ProductionCodeChanged': False}
    (Path(__file__).parent / 'Ready30NamedAnchorSelftest01.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'Status': proof['Status'], 'Checks': len(checks)}))


if __name__ == '__main__':
    TestNamedAnchors()
