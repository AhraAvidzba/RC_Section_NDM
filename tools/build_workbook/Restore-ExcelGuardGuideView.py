"""Возвращает видимость окна и свёрнутых строк после COM SaveCopyAs/AutoFit."""
from pathlib import Path
import sys
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import re

directory = Path(sys.argv[1])
candidate = directory / 'Candidate.xlsm'
ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with ZipFile(directory / 'Before.xlsm') as initial, ZipFile(directory / 'LiveBefore.xlsm') as live, ZipFile(candidate) as current:
    base_view = ET.fromstring(initial.read('xl/workbook.xml')).find(ns + 'bookViews')[0].get('visibility')
    entries = [(item, current.read(item.filename)) for item in current.infolist()]
    old_rows = {item.filename: {row.get('r'): row.attrib for row in ET.fromstring(live.read(item.filename)).find(ns + 'sheetData')}
                for item in current.infolist() if re.fullmatch(r'xl/worksheets/sheet\d+\.xml', item.filename)}
updated = []
for item, payload in entries:
    if item.filename == 'xl/workbook.xml':
        text = payload.decode('utf-8')
        def view(match):
            value = re.sub(r' visibility="[^"]*"', '', match[0])
            if base_view:
                value = value.replace('<workbookView ', '<workbookView visibility="' + base_view + '" ', 1)
            return value
        payload = re.sub(r'<workbookView\b[^>]*>', view, text).encode('utf-8')
    elif item.filename in old_rows:
        text = payload.decode('utf-8')
        def row(match):
            value = match[0]
            number = re.search(r'\br="(\d+)"', value)[1]
            original = old_rows[item.filename].get(number, {})
            for attribute in ('hidden', 'collapsed'):
                value = re.sub(r' ' + attribute + r'="[^"]*"', '', value)
                if attribute in original:
                    value = value.replace('<row ', '<row ' + attribute + '="' + original[attribute] + '" ', 1)
            return value
        payload = re.sub(r'<row\b[^>]*>', row, text).encode('utf-8')
    updated.append((item, payload))
temporary = candidate.with_suffix('.view.tmp')
with ZipFile(temporary, 'w') as archive:
    for item, payload in updated:
        archive.writestr(item, payload)
temporary.replace(candidate)
print('CANDIDATE_VIEW_RESTORED: original window visibility and row outline state')
