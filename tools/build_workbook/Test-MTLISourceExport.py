"""ДЛЯ ТЕСТОВ: сверка нативного экспорта экспериментальной книги с src/tests.

Используется штатная сверка VBE: регистр кода, печать Double и ANSI
комментариев. Строковые литералы, пробелы и типы целых констант сохраняются.
"""
import argparse
import hashlib
import json
import os
import posixpath
import re
import subprocess
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile


def sha(path):
    """Хеш проверяемого артефакта без его изменения."""
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def verify(root, export, book, equality):
    """Сверяет каждый импортируемый компонент и сохранённые значения настроек."""
    text = export.read_text(encoding='utf-8-sig')
    components = re.findall(r'(?m)^COMPONENT: ([^\n]+)$', text)
    compared = {v['Name']: v for v in equality['Components']}
    checks = []
    for directory in ['src', 'tests']:
        for path in sorted((root / directory).rglob('*')):
            if path.suffix not in {'.cls', '.bas'}:
                continue
            source = path.read_text(encoding='utf-8-sig')
            name = re.search(r'^Attribute VB_Name = "([^"]+)"', source, re.M)[1]
            check = compared.get(name, {})
            matches = check.get('Equal') is True and check.get('SourceSHA256') == sha(path)
            checks.append(dict(Path=path.relative_to(root).as_posix(), Component=name, SHA256=sha(path), Matches=matches, RawMatches=check.get('RawEqual')))
    with ZipFile(book) as archive:
        ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        named = next((v.text for v in workbook.findall('x:definedNames/x:definedName', ns) if v.attrib['name'] == 'rngExperimentalNumericsSettings'), None)
        config = next(v for v in workbook.findall('x:sheets/x:sheet', ns) if v.attrib['name'] == 'Config')
        relation_id = config.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
        relations = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        target = next(v.attrib['Target'] for v in relations if v.attrib['Id'] == relation_id)
        sheet_path = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))
        sheet = ET.fromstring(archive.read(sheet_path))
        strings = []
        if 'xl/sharedStrings.xml' in archive.namelist():
            strings = [''.join(v.itertext()) for v in ET.fromstring(archive.read('xl/sharedStrings.xml'))]
        cells = {v.attrib['r']: v for v in sheet.findall('x:sheetData/x:row/x:c', ns)}
        def value(address):
            cell = cells[address]
            v = cell.find('x:v', ns)
            if cell.attrib.get('t') == 's':
                return strings[int(v.text)]
            if cell.attrib.get('t') == 'inlineStr':
                return ''.join(cell.find('x:is', ns).itertext())
            return v.text if v is not None else ''
        validations = {v.attrib['sqref']: v for v in sheet.findall('x:dataValidations/x:dataValidation', ns)}
        # Addresses follow the exact range declared by this experiment's layout.
        row = int(re.search(r'\$A\$(\d+):', named)[1])
        method_address, tolerance_address = f'B{row+1}', f'B{row+2}'
        method_rule = validations[method_address]
        settings = dict(Method=value(method_address), RelativeTolerance=value(tolerance_address), ListFormula=method_rule.find('x:formula1', ns).text, NativeListAcceptance='Checked separately in saved/reopened Excel')
        settings['DefaultsValid'] = settings['Method'] == 'Discrete' and float(settings['RelativeTolerance']) == 0
        settings['RuleValid'] = settings['ListFormula'] == '"Discrete,MTLI"' and method_rule.attrib.get('showErrorMessage') == '1'
        styles = ET.fromstring(archive.read('xl/styles.xml'))
        xfs = styles.find('x:cellXfs', ns)
        borders = styles.find('x:borders', ns)
        fills = styles.find('x:fills', ns)
        frame = []
        for r in range(row, row+3):
            for col in 'ABCDE':
                xf = xfs[int(cells[f'{col}{r}'].attrib.get('s', 0))]
                border = borders[int(xf.attrib['borderId'])]
                frame.append(all(border.find('x:'+edge, ns) is not None and border.find('x:'+edge, ns).attrib.get('style') not in (None, 'none') for edge in ['left','right','top','bottom']))
        settings['FullBorders'] = all(frame)
        colors = []
        for address in [method_address, tolerance_address]:
            xf = xfs[int(cells[address].attrib.get('s', 0))]
            color = fills[int(xf.attrib['fillId'])].find('x:patternFill/x:fgColor', ns)
            colors.append(color.attrib.get('rgb') if color is not None else None)
        settings['InputColors'] = colors
        settings['InputFillValid'] = all(c in ('FFFFF2E6', '00FFF2E6') for c in colors)
    protected = [
        ('workbook/output/RC_Section_NDM.xlsm', 'CA57445B6AFFC701041B399ABBB57E16EFDFF5703BF02B568BA4EBA4200BC1D2'),
        ('workbook/output/VBA_All_Code.txt', '103F9E10E4144FEB13A0EF3A7B8FA77804EA5E0F2FE7D2BCF0D5A28A822092A8'),
    ]
    release = [dict(Path=p, SHA256=sha(root / p), Unchanged=sha(root / p) == expected) for p, expected in protected]
    accepted = equality['Passed'] and equality['ExportSHA256'] == sha(export) and checks and all(c['Matches'] for c in checks) and all(c['Unchanged'] for c in release) and named and "Config" in named and '$A$160:$E$162' in named
    accepted = accepted and all(settings[key] for key in ['DefaultsValid','RuleValid','FullBorders','InputFillValid'])
    return dict(Status='PASS' if accepted else 'FAIL', Components=len(components), SourceComponents=len(checks), Checks=checks, ExperimentalRange=named, Settings=settings, WorkbookSHA256=sha(book), ExportSHA256=sha(export), ProtectedRelease=release)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('export', type=Path)
    parser.add_argument('book', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    equality_path = args.output.with_name('SourceEquality.json')
    command = ['C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', 'tools/build_workbook/Verify-PostAudit03VbaEquality.ps1', '-ExportPath', str(args.export), '-ReportPath', str(equality_path)]
    native_env = {k: v for k, v in os.environ.items() if k.upper() != 'PSMODULEPATH'}
    native_env['PSModulePath'] = os.path.join(os.environ.get('SystemRoot', 'C:/Windows'), 'System32', 'WindowsPowerShell', 'v1.0', 'Modules')
    completed = subprocess.run(command, check=False, capture_output=True, env=native_env)
    print(completed.stdout.decode('utf-8', errors='replace'))
    if completed.returncode:
        print(completed.stderr.decode('utf-8', errors='replace'))
        raise SystemExit('Native source equality failed; primary report preserved')
    equality = json.loads(equality_path.read_text(encoding='utf-8-sig'))
    result = verify(Path.cwd(), args.export, args.book, equality)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(result['Status'], result['SourceComponents'], 'source components')
    if result['Status'] != 'PASS':
        print('Mismatches:', [c['Path'] for c in result['Checks'] if not c['Matches']])
    raise SystemExit(result['Status'] != 'PASS')
