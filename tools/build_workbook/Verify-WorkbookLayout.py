"""Проверяет сохранность книги и отсутствие переоформления входных листов при расчёте."""
from __future__ import annotations

import argparse
import importlib
import hashlib
import json
import re
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

_reader = importlib.import_module('Verify-SectionRecipeGuide')
read, canonical, S = _reader.read, _reader.canonical, _reader.S


def compare_styles(before, after):
    failures, count = [], 0
    row_fields = {'ht', 'hidden', 'customHeight', 's', 'customFormat', 'outlineLevel', 'collapsed', 'thickTop', 'thickBot'}
    for name, old in before['sheets'].items():
        if name == 'Results':
            continue
        new = after['sheets'][name]
        for address in old['cells'].keys() | new['cells'].keys():
            count += 1
            if old['cells'].get(address, {}).get('style') != new['cells'].get(address, {}).get('style'):
                failures.append(f'{name}!{address}: cell formatting changed')
        for row in old['rows'].keys() | new['rows'].keys():
            count += 1
            a = {k:v for k,v in old['rows'].get(row, {}).items() if k in row_fields}
            b = {k:v for k,v in new['rows'].get(row, {}).items() if k in row_fields}
            if a != b:
                failures.append(f'{name}!row {row}: row formatting changed')
        count += 1
        if old['columns'] != new['columns']:
            failures.append(f'{name}: column dimensions/formatting changed')
        for tag in ('mergeCells', 'conditionalFormatting', 'sheetFormatPr'):
            count += 1
            if [canonical(x) for x in old['document'].findall(S+tag)] != [canonical(x) for x in new['document'].findall(S+tag)]:
                failures.append(f'{name}: {tag} changed')
    return count, failures


def verify_runtime(directory):
    """Проверяет самостоятельный runtime-прогон, не подменяя отдельную приёмку пересборки."""
    manifest = json.loads((directory/'Run.json').read_text('utf-8-sig'))
    failures = []
    if manifest['Status'] != 'RUNTIME_COMPLETED':
        failures.append('Runtime did not finish')
    if hashlib.sha256((directory/'Runtime.xlsm').read_bytes()).hexdigest().upper() != manifest['SourceWorkbookSHA256']:
        failures.append('Private source differs from the runtime manifest')
    baseline = read(directory/'BeforeFormation.xlsm')
    stages = []
    for stage in ('BeforeCalculation', 'AfterCalculation1', 'AfterPlot1', 'AfterCalculation2', 'AfterPlot2'):
        count, errors = compare_styles(baseline, read(directory/f'{stage}.xlsm'))
        failures.extend(f'{stage}: {error}' for error in errors)
        stages.append(dict(Stage=stage, Checks=count, Failures=errors))
    result = dict(Status='FAIL' if failures else 'PASS', SourceSHA256=manifest['SourceWorkbookSHA256'],
                  Scope='Input formatting through formation, calculation and refresh',
                  Stages=stages, Failures=failures)
    (directory/'RuntimeFormatting.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps(result, ensure_ascii=False))
    if failures:
        raise SystemExit(1)


def verify(directory):
    old, new = read(directory/'Saved.xlsm'), read(directory/'Candidate.xlsm')
    failures, checked, cached = [], 0, []
    headers = {2, 21, 59, 93}
    for name, sheet in old['sheets'].items():
        actual = new['sheets'][name]
        for address in sheet['cells'].keys() | actual['cells'].keys():
            a, b = sheet['cells'].get(address), actual['cells'].get(address)
            checked += 1
            is_header = name == 'Сечения' and int(''.join(filter(str.isdigit,address))) in headers
            if a is None or b is None:
                if is_header and a is None and b is not None and b['formula'] is None and b['value'] is None:
                    continue
                failures.append(f'Preservation {name}!{address}: cell added/removed')
                continue
            if a['formula'] != b['formula'] or (a['formula'] is None and a['value'] != b['value']):
                failures.append(f'Preservation {name}!{address}: value/formula changed')
            if a['formula'] is not None and a['value'] != b['value']:
                cached.append(f'{name}!{address}')
            if not is_header:
                if a['style'] != b['style']:
                    failures.append(f'Preservation {name}!{address}: unexpected formatting change')
        old_rows = {k:v for k,v in sheet['rows'].items() if name != 'Сечения' or k not in headers}
        new_rows = {k:v for k,v in actual['rows'].items() if name != 'Сечения' or k not in headers}
        if old_rows != new_rows:
            failures.append(f'Preservation {name}: row metadata changed')
        for tag in ('cols', 'mergeCells', 'hyperlinks', 'dataValidations', 'conditionalFormatting'):
            if [canonical(x) for x in sheet['document'].findall(S+tag)] != [canonical(x) for x in actual['document'].findall(S+tag)]:
                failures.append(f'Preservation {name}: {tag} changed')
    if old['names'] != new['names']:
        failures.append('Preservation: named ranges changed')
    if old['drawings'] != new['drawings']:
        failures.append('Preservation: drawing XML changed')

    stages=[]
    runtime=json.loads((directory/'Runtime/Run.json').read_text('utf-8-sig'))
    candidate_hash=hashlib.sha256((directory/'Candidate.xlsm').read_bytes()).hexdigest().upper()
    if runtime['Status'] != 'RUNTIME_COMPLETED' or runtime['SourceWorkbookSHA256'] != candidate_hash:
        failures.append('Runtime evidence does not belong to the accepted candidate')
    baseline=read(directory/'Runtime/BeforeCalculation.xlsm')
    for stage in ('AfterCalculation1','AfterPlot1','AfterCalculation2','AfterPlot2'):
        count, errors=compare_styles(baseline, read(directory/'Runtime'/f'{stage}.xlsm'))
        failures.extend(f'{stage}: {error}' for error in errors)
        stages.append({'Stage':stage,'Checks':count,'Failures':errors})

    # Полный поиск команд оформления в рабочих исходниках и модулях книги/листов.
    pattern=re.compile(r'RowHeight|ColumnWidth|AutoFit|StandardHeight|StandardWidth|ClearFormats|\.Interior\.|\.Font\.|\.NumberFormat|\.HorizontalAlignment|\.VerticalAlignment|\.WrapText|\.Borders\b',re.I)
    expected={'CBatchResultWriter','CCrackSummaryWriter','CNDMResultsWriter','CStabilitySummaryWriter','CStrengthSummaryWriter','CSectionPlotter','modStatusFormatting','CCalculationSnapshotStore'}
    source=Path(__file__).resolve().parents[2]/'src'
    format_owners=[]
    for file in sorted(source.rglob('*')):
        if file.suffix not in {'.bas','.cls'}:
            continue
        matches=[i for i,line in enumerate(file.read_text('utf-8-sig').splitlines(),1) if not line.lstrip().startswith("'") and pattern.search(line)]
        if matches:
            format_owners.append({'File':file.relative_to(source).as_posix(),'Lines':matches})
            if file.stem not in expected:
                failures.append(f'Unexpected runtime formatting owner: {file.name}')
    exported=(directory/'CandidateNative.txt').read_text('utf-8-sig')
    documents=[]
    for match in re.finditer(r'(?ms)^COMPONENT: ([^\n]+)\nTYPE: 100\nLINES: \d+\n=+\n(.*?)(?=\n=+\nCOMPONENT: |\Z)',exported):
        documents.append(match[1])
        if any(pattern.search(line) for line in match[2].splitlines() if not line.lstrip().startswith("'")):
            failures.append(f'Unexpected document formatting: {match[1]}')
    if len(documents)!=6:
        failures.append(f'Document modules not fully audited: {len(documents)}')

    fresh=read(directory/'RebuiltFinal.xlsm')
    with ZipFile(directory/'RebuiltFinal.xlsm') as archive:
        for part in archive.namelist():
            if not part.endswith('.xml') or not part.startswith(('xl/worksheets/','xl/charts/')):
                continue
            document=ET.fromstring(archive.read(part))
            for node in document.iter():
                if node.tag.split('}')[-1] == 'f' and ('#REF!' in (node.text or '') or '__NDMLayoutBuild' in (node.text or '')):
                    failures.append(f'Fresh build: broken formula {part}: {node.text}')
                if node.tag == S+'c' and node.get('t') == 'e':
                    failures.append(f'Fresh build: cached Excel error {part}!{node.get("r")}: {node.findtext(S+"v")}')
    for name, sheet in fresh['sheets'].items():
        for link in sheet['document'].findall(S+'hyperlinks/'+S+'hyperlink'):
            target=link.get('location','')
            if '#REF!' in target or '__NDMLayoutBuild' in target:
                failures.append(f'Fresh build: broken hyperlink {name}!{link.get("ref")}')
    result={'Status':'FAIL' if failures else 'PASS','PreservedCells':checked,'RecalculatedFormulaCaches':cached,
            'RuntimeStages':stages,'RuntimeFormattingOwners':format_owners,'AuditedDocumentModules':documents,
            'FreshFormulaErrors':0 if not any('Fresh build' in e for e in failures) else None,'Failures':failures}
    (directory/'PreservationAndRuntime.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({k:v for k,v in result.items() if k not in {'RuntimeStages','RuntimeFormattingOwners'}},ensure_ascii=False))
    if failures:
        raise SystemExit(1)


if __name__ == '__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--directory', type=Path, required=True)
    parser.add_argument('--runtime-only', action='store_true')
    arguments = parser.parse_args()
    (verify_runtime if arguments.runtime_only else verify)(arguments.directory)
