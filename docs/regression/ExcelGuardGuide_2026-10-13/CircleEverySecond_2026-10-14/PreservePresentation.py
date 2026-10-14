"""Сохраняет книгу до/после миграции Config; код и два новых поля берёт из кандидата."""
import copy
import json
import re
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as ET

directory=Path(__file__).resolve().parent
S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
def canonical(e):
    return e.tag,tuple(sorted(e.attrib.items())),e.text,tuple(canonical(c) for c in e)
def xml(e):
    return ET.tostring(e,encoding='utf-8',xml_declaration=True)

with ZipFile(directory/'Baseline.xlsm') as old, ZipFile(directory/'Candidate.xlsm') as new:
    result={name:old.read(name) for name in old.namelist()}
    book=ET.fromstring(old.read('xl/workbook.xml'))
    current_book=ET.fromstring(new.read('xl/workbook.xml'))
    relations={e.get('Id'):e.get('Target') for e in ET.fromstring(old.read('xl/_rels/workbook.xml.rels'))}
    sheet_paths={e.get('name'):relations[e.get(R+'id')] for e in book.find(S+'sheets')}
    sheet_paths={name:(path.lstrip('/') if path.startswith('/') else 'xl/'+path) for name,path in sheet_paths.items()}
    names=book.find(S+'definedNames')
    current_names={e.get('name'):e for e in current_book.find(S+'definedNames')}
    changed_names=('rngCircleGeometry','rngRectSetGeometry','rngRoundedRectangleGeometry','rngHollowRectangleGeometry')
    for e in names:
        if e.get('name') in changed_names:
            e.text=current_names[e.get('name')].text
    for name in ('rngCircleBind2Choices','rngCircleBind3Choices'):
        names.append(copy.deepcopy(current_names[name]))
    result['xl/workbook.xml']=xml(book)
    result['xl/vbaProject.bin']=new.read('xl/vbaProject.bin')

    # Excel переупорядочивает одинаковые cellXfs. Используем его валидный styles.xml,
    # переводя исходные индексы без изменения самих шрифтов, рамок или выравнивания.
    styles=ET.fromstring(old.read('xl/styles.xml'))
    current_styles=ET.fromstring(new.read('xl/styles.xml'))
    for group in ('numFmts','fonts','fills','borders','cellStyleXfs'):
        assert all(canonical(a)==canonical(b) for a,b in zip(styles.find(S+group),current_styles.find(S+group))),group
    native_xfs={canonical(e):i for i,e in enumerate(current_styles.find(S+'cellXfs'))}
    old_to_native={i:native_xfs[canonical(e)] for i,e in enumerate(styles.find(S+'cellXfs'))}
    result['xl/styles.xml']=new.read('xl/styles.xml')
    result['xl/calcChain.xml']=new.read('xl/calcChain.xml')
    def translate_styles(document):
        for element in document.iter():
            attribute='style' if element.tag==S+'col' else 's' if element.tag in (S+'c',S+'row') else None
            if attribute and attribute in element.attrib:
                element.set(attribute,str(old_to_native[int(element.get(attribute))]))

    old_config=ET.fromstring(old.read(sheet_paths['Config']))
    config=ET.fromstring(new.read(sheet_paths['Config']))
    old_cells={e.get('r'):e for e in old_config.iter(S+'c')}
    cells={e.get('r'):e for e in config.iter(S+'c')}
    strings=[''.join(e.itertext()) for e in ET.fromstring(new.read('xl/sharedStrings.xml'))]
    def moved(address):
        col,raw=re.fullmatch(r'([A-Z]+)(\d+)',address).groups();row=int(raw)
        return col+str(row+2 if 'H'<=col<='P' and len(col)==1 and row>=96 else row)
    formula_adjustments=[]
    for address,source in old_cells.items():
        destination=moved(address)
        if destination not in cells:
            assert source.find(S+'v') is None and source.find(S+'f') is None and source.find(S+'is') is None,address
            continue
        target=cells[destination]
        formula=copy.deepcopy(target.find(S+'f'))
        if source.find(S+'f') is not None:
            if canonical(source.find(S+'f')) != (canonical(formula) if formula is not None else None):
                formula_adjustments.append(destination)
        target.attrib.clear();target.attrib.update(source.attrib);target.set('r',destination)
        if 's' in target.attrib:target.set('s',str(old_to_native[int(target.get('s'))]))
        for child in list(target):target.remove(child)
        for child in source:target.append(copy.deepcopy(child))
        if source.find(S+'f') is not None and formula is not None:
            target.replace(target.find(S+'f'),formula)
    mapped_old={moved(a) for a in old_cells}
    for address,cell in cells.items():
        if address not in mapped_old:
            if cell.get('t')=='s':
                value=strings[int(cell.find(S+'v').text)]
                cell.remove(cell.find(S+'v'));cell.set('t','inlineStr')
                ET.SubElement(ET.SubElement(cell,S+'is'),S+'t').text=value
    result[sheet_paths['Config']]=xml(config)

    help_doc=ET.fromstring(old.read(sheet_paths['Справка']))
    old_strings=[''.join(e.itertext()) for e in ET.fromstring(old.read('xl/sharedStrings.xml'))]
    changed_help=[]
    addition=' Rebar.Bind2row и Rebar.Bind3row выбираются независимо: EachBar — у каждого стержня первого ряда; EverySecondBar — у первого, третьего и следующих нечётных. По умолчанию оба ряда используют EachBar.'
    for cell in help_doc.iter(S+'c'):
        value=cell.find(S+'v')
        text=''.join(cell.find(S+'is').itertext()) if cell.get('t')=='inlineStr' else old_strings[int(value.text)] if cell.get('t')=='s' and value is not None else ''
        if text.startswith('Rebar.Loc2row и Rebar.Loc3row задают положение дополнительных стержней:'):
            assert 'Rebar.Bind2row' not in text
            for child in list(cell):cell.remove(child)
            cell.set('t','inlineStr');ET.SubElement(ET.SubElement(cell,S+'is'),S+'t').text=text+addition
            changed_help.append(cell.get('r'))
    assert len(changed_help)==1,changed_help
    translate_styles(help_doc)
    result[sheet_paths['Справка']]=xml(help_doc)

    # Формулы обычных листов могли автоматически получить новые адреса Config.
    for sheet,path in sheet_paths.items():
        if sheet in ('Config','Справка'):continue
        previous=ET.fromstring(old.read(path));current=ET.fromstring(new.read(path))
        actual={e.get('r'):e for e in current.iter(S+'c')}
        changed=False
        for cell in previous.iter(S+'c'):
            f=cell.find(S+'f');g=actual.get(cell.get('r'))
            g=g.find(S+'f') if g is not None else None
            if f is not None and g is not None and canonical(f)!=canonical(g):
                assert 'Config!' in (f.text or '') or "'Config'!" in (f.text or ''),(sheet,cell.get('r'),f.text,g.text)
                cell.replace(f,copy.deepcopy(g));formula_adjustments.append(sheet+'!'+cell.get('r'));changed=True
        translate_styles(previous)
        result[path]=xml(previous)
    temp=directory/'Candidate.preserved.tmp'
    with ZipFile(temp,'w',ZIP_DEFLATED) as out:
        for entry in old.infolist():out.writestr(entry,result[entry.filename])
temp.replace(directory/'Candidate.xlsm')
(directory/'PresentationPreserved.json').write_text(json.dumps({'Passed':True,'HelpCellsChanged':changed_help,'ConfigFormulaAdjustments':formula_adjustments,'UnrelatedZIPPartsRestored':True,'ExistingStylesRetained':True},ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print('PRESENTATION_PRESERVED',changed_help,'FORMULA_ADJUSTMENTS',formula_adjustments)
