"""Возвращает прежнее оформление листов; оставляет только новый абзац Справки."""
import json
import re
from pathlib import Path
from lxml import etree as ET
from zipfile import ZipFile, ZIP_DEFLATED

directory = Path(__file__).resolve().parent
root = directory.parents[3]
baseline, candidate = directory/'Baseline.xlsm', directory/'Candidate.xlsm'
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
layout = json.loads((root/'docs/geometry_refactor/SECTION_RECIPE_GUIDE_LAYOUT.json').read_text(encoding='utf-8'))

with ZipFile(baseline) as old, ZipFile(candidate) as new:
    workbook = ET.fromstring(old.read('xl/workbook.xml'))
    names = {e.get('name'): e.text for e in workbook.find(S+'definedNames')}
    first = int(re.findall(r'\$A\$(\d+)', names['rngSectionRecipeGuide'])[0])
    relationships = {e.get('Id'): e.get('Target') for e in ET.fromstring(old.read('xl/_rels/workbook.xml.rels'))}
    sheet = next(e for e in workbook.find(S+'sheets') if e.get('name') == 'Справка')
    target = relationships[sheet.get(R+'id')]
    target = target.lstrip('/') if target.startswith('/') else 'xl/'+target
    document = ET.fromstring(old.read(target))
    new_document = ET.fromstring(new.read(target))
    old_cells = {e.get('r'): e for e in document.iter(S+'c')}
    new_rows = {int(e.get('r')): e for e in new_document.iter(S+'row')}
    old_rows = {int(e.get('r')): e for e in document.iter(S+'row')}
    shared_document = ET.fromstring(old.read('xl/sharedStrings.xml'))
    strings = [''.join(e.itertext()) for e in shared_document]
    obsolete_strings = []
    for index, value in enumerate(strings):
        if 'UpdateSectionRecipeConfig' in value:
            item = shared_document[index]
            for child in list(item): item.remove(child)
            ET.SubElement(item, S+'t').text = value.replace('UpdateSectionRecipeConfig', 'UpdateSectionRecipeTypes')
            obsolete_strings.append(index)
        elif 'SetPublication' in value:
            replacement_row = next(row['text'] for row in layout['rows'] if 'AnnotationGroup' in row['text'] and 'ConfigureDependentContact' in row['text'])
            item = shared_document[index]
            for child in list(item):
                item.remove(child)
            ET.SubElement(item, S+'t').text = replacement_row
            obsolete_strings.append(index)
    changes = []
    for i, row in enumerate(layout['rows']):
        number = first+i
        cell = old_cells['A'+str(number)]
        v = cell.find(S+'v')
        previous = ''.join(cell.find(S+'is').itertext()) if cell.get('t')=='inlineStr' else strings[int(v.text)] if cell.get('t') == 's' and v is not None else (v.text if v is not None else '')
        if (previous or '') == row['text']:
            continue
        for child in list(cell):
            if child.tag in (S+'v', S+'is', S+'f'):
                cell.remove(child)
        cell.set('t', 'inlineStr')
        text = ET.SubElement(ET.SubElement(cell, S+'is'), S+'t')
        text.text = row['text']
        for key in ('ht', 'customHeight'):
            if key in new_rows[number].attrib:
                old_rows[number].set(key, new_rows[number].get(key))
        changes.append('A'+str(number))
    replacement = {name: old.read(name) for name in old.namelist()}
    # Точечные пояснения согласованного коэффициента в существующих ячейках.
    edits = json.loads((directory/'PresentationText.json').read_text(encoding='utf-8'))
    changed_text = []
    for sheet_name in ('Config', 'Справка'):
        sheet_info = next(e for e in workbook.find(S+'sheets') if e.get('name') == sheet_name)
        sheet_target = relationships[sheet_info.get(R+'id')]
        sheet_target = sheet_target.lstrip('/') if sheet_target.startswith('/') else 'xl/'+sheet_target
        sheet_document = document if sheet_name == 'Справка' else ET.fromstring(old.read(sheet_target))
        cells = {e.get('r'): e for e in sheet_document.iter(S+'c')}
        for key, value in edits.items():
            name, address = key.split('!')
            if name != sheet_name: continue
            cell = cells[address]
            for child in list(cell):
                if child.tag in (S+'v', S+'is', S+'f'): cell.remove(child)
            cell.set('t', 'inlineStr')
            ET.SubElement(ET.SubElement(cell, S+'is'), S+'t').text = value
            changed_text.append(key)
        replacement[sheet_target] = ET.tostring(sheet_document, encoding='utf-8', xml_declaration=True)
    replacement['xl/vbaProject.bin']=new.read('xl/vbaProject.bin')
    replacement[target] = ET.tostring(document, encoding='utf-8', xml_declaration=True)
    if obsolete_strings:
        replacement['xl/sharedStrings.xml'] = ET.tostring(shared_document, encoding='utf-8', xml_declaration=True)
    temp = candidate.with_suffix('.preserved.tmp')
    with ZipFile(temp, 'w', ZIP_DEFLATED) as out:
        for entry in old.infolist():
            out.writestr(entry, replacement[entry.filename])
temp.replace(candidate)
(directory/'PresentationPreserved.json').write_text(json.dumps({'Passed': True, 'HelpTextChanged': changes, 'ObsoleteSharedStringEntriesReplaced': obsolete_strings, 'OtherWorksheetXMLAndStylesFromBaseline': True, 'VBAProjectUntouched': True}, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
print('PRESENTATION_PRESERVED:', changes)
