"""ДЛЯ ТЕСТОВ: независимая проверка нагрузок, снимков и прежних результатов.

Читает OOXML и двоичные снимки без Excel/VBA. Нативные расчёты выполняются
отдельным Test-BatchSummaryLoads.ps1 на частных копиях готовой книги.
"""
import argparse
import hashlib
import importlib.util
import json
import math
import re
import struct
from functools import lru_cache
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as E


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestModule(name):
    """Подключает существующие независимые читатели с точными типами значений."""
    spec = importlib.util.spec_from_file_location(name.replace('-', '_'), Path(__file__).with_name(name + '.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


XML = TestModule('Verify-MultipleSnapshotFiles')
BINARY = TestModule('Compare-SnapshotPayloads')
GUIDE = TestModule('Verify-SectionRecipeGuide')


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestAddress(row, column):
    """Преобразует координаты в A1 без использования библиотеки Excel."""
    label = ''
    while column:
        column, digit = divmod(column - 1, 26)
        label = chr(65 + digit) + label
    return label + str(row)


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestPosition(address):
    """Разбирает единственный адрес, чтобы не привязывать тест к размещению ввода."""
    match = re.fullmatch(r'([A-Z]+)([0-9]+)', address.replace('$', ''))
    column = 0
    for letter in match[1]:
        column = column * 26 + ord(letter) - 64
    return int(match[2]), column


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestValue(data, sheet, address):
    """Возвращает исходный Double, строку либо пустое значение OOXML."""
    value = data[sheet]['values'].get(address)
    if value is None:
        return None
    return struct.unpack('<d', bytes.fromhex(value[1]))[0] if value[0] == 'n' else value[1]


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestAnchor(path, name):
    """Получает положение таблицы из реального имени сохранённой книги."""
    with ZipFile(path) as archive:
        document = E.fromstring(archive.read('xl/workbook.xml'))
    ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    node = next(n for n in document.findall('m:definedNames/m:definedName', ns) if n.get('name') == name)
    sheet, address = node.text.split('!', 1)
    row, column = TestPosition(address.split(':')[0])
    return sheet.strip("'"), row, column


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestClose(actual, expected):
    """Учитывает лишь погрешность конвертации единиц, без инженерного допуска."""
    assert isinstance(actual, (int, float)) and math.isclose(actual, expected, rel_tol=5e-15, abs_tol=1e-12), (actual, expected)


# -------------------- ДЛЯ ТЕСТОВ --------------------
@lru_cache(maxsize=32)
def TestBook(path):
    """Кэширует только независимое чтение неизменных проверяемых файлов."""
    return XML.TestReadBook(path)


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestCase(directory, case):
    """Проверяет исходный ввод, преобразование единиц, рабочие нагрузки и Worst."""
    path = directory / case['Workbook']
    sheets, _ = TestBook(path)
    source_sheet, source_row, source_column = TestAnchor(path, 'rngLoadCombinations')
    unit_sheet, unit_row, unit_column = TestAnchor(path, 'rngUnitSettings')
    input_units = {}
    for offset in range(7):
        key = TestValue(sheets, unit_sheet, TestAddress(unit_row + offset, unit_column))
        input_units[key] = TestValue(sheets, unit_sheet, TestAddress(unit_row + offset, unit_column + 1))
    force_factors = {'N': 1., 'kN': 1000., 'tf': 9806.65}
    moment_factors = {'N*mm': 1., 'kN*m': 1000000., 'tf*m': 9806650.}
    factors = [force_factors[input_units['Force']] / force_factors[case['ForceUnit']],
               moment_factors[input_units['Moment']] / moment_factors[case['MomentUnit']]]
    original_checks = working_checks = 0
    amplified = filtered = False
    stable_sheet, stable_row, stable_column = TestAnchor(path, 'rngStabilitySummaryAnchor')
    worst_row = None
    for row in range(13, 43):
        identifier = TestValue(sheets, 'Results', 'A' + str(row))
        if not identifier:
            continue
        source = next(r for r in range(source_row + 1, source_row + 31)
                      if TestValue(sheets, source_sheet, TestAddress(r, source_column)) == identifier)
        for offset in range(3):
            raw = TestValue(sheets, source_sheet, TestAddress(source, source_column + 1 + offset))
            expected = raw * factors[0 if offset == 0 else 1]
            TestClose(TestValue(sheets, 'Results', TestAddress(row, 28 + offset)), expected)
            original_checks += 1
        stable = next(r for r in range(stable_row, stable_row + 31)
                      if TestValue(sheets, stable_sheet, TestAddress(r, stable_column)) == identifier)
        status = TestValue(sheets, stable_sheet, TestAddress(stable, stable_column + 2))
        if status == 'OK':
            for offset in range(3):
                TestClose(TestValue(sheets, 'Results', TestAddress(row, 31 + offset)),
                          TestValue(sheets, stable_sheet, TestAddress(stable, stable_column + 11 + offset)))
                working_checks += 1
            TestClose(TestValue(sheets, 'Results', 'AB' + str(row)), TestValue(sheets, 'Results', 'AE' + str(row)))
            amplified |= any(abs(TestValue(sheets, 'Results', TestAddress(row, c + 3))) > abs(TestValue(sheets, 'Results', TestAddress(row, c))) * 1.01
                             for c in (29, 30))
        if case['Filtered']:
            filtered |= any(TestValue(sheets, 'Results', TestAddress(row, c)) != 0 and TestValue(sheets, 'Results', TestAddress(row, c + 3)) == 0 for c in (29, 30))
        if identifier == case['WorstID']:
            worst_row = row
    assert original_checks == case['Count'] * 3
    assert worst_row
    for column in range(28, 34):
        assert sheets['Results']['values'].get(TestAddress(12, column)) == sheets['Results']['values'].get(TestAddress(worst_row, column))
    if case['Name'] == 'File_Minimal':
        assert amplified, 'Fixture must exercise actual amplified moments'
    if case['Filtered']:
        assert filtered, 'Fixture must retain a raw moment removed from working loads'
    if case['Storage'] == 'File':
        _, _, tables = BINARY.TestSnapshotReader(directory / (case['Name'] + '.ndm')).TestRead()
        assert len(tables) == 16 and len(tables['Summary'][1]) == 33
    return dict(Name=case['Name'], OriginalLoadChecks=original_checks, WorkingLoadChecks=working_checks,
                WorstID=case['WorstID'], AmplifiedMoments=amplified, FilteredOriginalPreserved=filtered,
                WorkbookSHA256=hashlib.sha256(path.read_bytes()).hexdigest().upper())


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestRestores(directory, native):
    """Сверяет все Results и реальные фигуры; остальные листы должны быть неизменны."""
    records = []
    for item in native['Restores']:
        before, _ = TestBook(directory / item['Before'])
        after, drawing = TestBook(directory / item['After'])
        expected, expected_drawing = TestBook(directory / item['ExpectedBook'])
        for sheet in before.keys() - {'Results'}:
            assert before[sheet] == after[sheet], ('Non-Results sheet changed', sheet, item['Name'])
        a, b = expected['Results']['values'], after['Results']['values']
        # При No расчёт сохраняет старую шапку подавленной File-таблицы элементов.
        # Загрузка файла очищает эту пустую таблицу вместе с её прежней шапкой.
        # Разрешаем только удаление десяти точно проверенных текстов, не данных.
        _, header_row, header_column = TestAnchor(directory / item['After'], 'rngNDMElementResults')
        labels = ['RunID', 'LoadCase', 'ProfileId', 'StateType', 'ElementID', 'MaterialType', 'Strain', 'Stress, MPa', 'PhysicalState']
        empty_headers = {TestAddress(header_row, header_column + offset): label for offset, label in enumerate(labels)}
        empty_headers[TestAddress(header_row - 1, header_column)] = 'Результаты элементов по сочетаниям'
        removed_headers = []
        differences = []
        for address in a.keys() | b.keys():
            if a.get(address) == b.get(address):
                continue
            if address in empty_headers and TestValue(expected, 'Results', address) == empty_headers[address] and address not in b:
                removed_headers.append(address)
            else:
                differences.append(address)
        assert not differences, (item['Name'], differences[:10])
        assert drawing == expected_drawing
        assert item['SolverCalls'] == 0
        records.append(dict(Name=item['Name'], ResultDataExact=True, OtherSheetsUnchanged=True,
                            ChartShapesExact=True, AfterReopenButton=bool(item.get('AfterReopenButton')), SolverCalls=0,
                            SuppressedEmptyTableHeadersRemoved=sorted(removed_headers)))
    before, _ = TestBook(directory / 'Legacy_Before.xlsm')
    after, _ = TestBook(directory / 'Legacy_After.xlsm')
    for sheet in before.keys() - {'Results'}:
        assert before[sheet] == after[sheet], ('Legacy restore changed another sheet', sheet)
    _, elapsed, tables = BINARY.TestSnapshotReader(directory / 'Legacy.ndm').TestRead()
    assert len(tables['Summary'][1]) == 26
    for row in range(12, 43):
        for column in range(28, 34):
            assert TestValue(after, 'Results', TestAddress(row, column)) is None
    assert TestValue(after, 'Results', 'AB11') == 'Не сохранено в этом снимке'
    for column, values in enumerate(tables['Summary'][1], 1):
        for row, scalar in enumerate(values, 1):
            expected = None if scalar[0] == 0 else scalar[1].decode('utf-16-le') if scalar[0] == 1 else struct.unpack('<d', scalar[1])[0]
            if row == 4 and column == 1:
                expected = struct.unpack('<d', elapsed)[0] # Окончательное время — поле контейнера, записанное после подготовки сводки.
            actual = TestValue(after, 'Results', TestAddress(row, column))
            if isinstance(expected, float):
                TestClose(actual, expected)
            else:
                assert actual == (expected or None), ('Legacy value changed', row, column)
    return records


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestBaseline(baseline, current, cases):
    """Сравнивает прежние результаты; исключает только новый блок, RunID и время."""
    checks = 0
    records = []
    for case in cases:
        old_path, new_path = baseline / case['Workbook'], current / case['Workbook']
        old, _ = TestBook(old_path)
        new, _ = TestBook(new_path)
        sh, r, c = TestAnchor(old_path, 'rngNDMSectionGeometry')
        old_id = old[sh]['values'][TestAddress(r + 1, c)]
        new_id = new[sh]['values'][TestAddress(r + 1, c)]
        excluded_ids = 0
        a, b = old['Results']['values'], new['Results']['values']
        for address in a.keys() | b.keys():
            row, column = TestPosition(address)
            if row <= 43 and column > 26 or address == 'A4':
                continue
            x, y = a.get(address), b.get(address)
            if x == old_id and y == new_id:
                excluded_ids += 1
                continue
            assert x == y, ('Existing result changed', case['Name'], address, x, y)
            checks += 1
        if case['Storage'] == 'File':
            ra, _, ta = BINARY.TestSnapshotReader(baseline / (case['Name'] + '.ndm')).TestRead()
            rb, _, tb = BINARY.TestSnapshotReader(current / (case['Name'] + '.ndm')).TestRead()
            for key, (height, columns) in ta.items():
                other_height, other = tb[key]
                assert height == other_height
                assert len(other) == (33 if key == 'Summary' else len(columns))
                for column, values in enumerate(columns):
                    for row, x in enumerate(values):
                        y = other[column][row]
                        if (key.startswith('rngNDM') or key.startswith('LoadCase.')) and column == 0 and row:
                            assert x == (2, ra) and y == (2, rb)
                            continue
                        if key == 'Summary' and column == 0 and row == 3:
                            continue
                        assert BINARY.normalized(x) == BINARY.normalized(y), ('Binary result changed', key, row, column)
                        checks += 1
        records.append(dict(Name=case['Name'], ExistingResultsExact=True, ExcludedRunIDCells=excluded_ids))
    return dict(Cases=records, ValuesCompared=checks, Exclusions=['New load block and moved legend', 'RunID', 'Measured elapsed time'])


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestInputs(before, after):
    """Проверяет неизменность всех входных чисел и формул при пересборке."""
    old, new = GUIDE.read(before), GUIDE.read(after)
    descriptions = []
    for sheet in ('Config', 'Сечения', 'Расчет'):
        a, b = old['sheets'][sheet]['cells'], new['sheets'][sheet]['cells']
        for address in a.keys() | b.keys():
            x, y = a.get(address, {}), b.get(address, {})
            if (x.get('value'), x.get('formula')) == (y.get('value'), y.get('formula')):
                continue
            assert sheet == 'Config' and address in ('D43', 'D89', 'D90', 'D91', 'D92'), ('Input changed', sheet, address)
            assert x.get('formula') is None and y.get('formula') is None
            descriptions.append(dict(Cell=address, Before=x.get('value'), After=y.get('value')))
    return dict(InputValuesAndFormulasUnchanged=True, StaticDescriptionsSynchronized=descriptions)


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestMain():
    """Требует обе нативные пары и отдельное сравнение с исходной книгой."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('directory', type=Path)
    args = parser.parse_args()
    base = args.directory
    result = dict(Status='PASS', Modes=[], Baseline=None)
    for mode in ('No', 'Yes'):
        directory = base / ('Native_' + mode + '_05')
        native = json.loads((directory / 'NativeSummaryLoads.json').read_text(encoding='utf-8-sig'))
        assert len(native['Cases']) == 4 and len(native['Restores']) == 4
        result['Modes'].append(dict(Mode=mode, Cases=[TestCase(directory, case) for case in native['Cases']],
                                    Restores=TestRestores(directory, native), LegacySnapshotCompatible=True,
                                    NativeChecks=native['Checks'], SourceSHA256=native['SourceSHA256']))
    result['Baseline'] = TestBaseline(base / 'Baseline_No_05', base / 'Native_No_05', native['Cases'])
    result['Inputs'] = TestInputs(base / 'Baseline.xlsm', base / 'Final/RC_Section_NDM.xlsm')
    (base / 'IndependentAcceptance.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(dict(Status='PASS', Modes=2, RealCalculations=8, Restores=8, LegacyRestores=2,
                         BaselineValuesCompared=result['Baseline']['ValuesCompared'])))


if __name__ == '__main__':
    TestMain()
