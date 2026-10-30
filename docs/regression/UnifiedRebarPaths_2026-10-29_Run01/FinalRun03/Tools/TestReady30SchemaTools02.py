"""ДЛЯ ТЕСТОВ: чистые OpenXML-операции будущего end-only обновления Ready30.

Не открывает Excel, не читает реальные книги при --selftest и ничего не публикует.
Старые таблицы стилей дополняются без изменения существующих индексов.
"""
from copy import deepcopy
from lxml import etree as ET
import argparse
import json
import re

S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS={'s':S}
def TestTag(name):
    """Возвращает полное имя элемента SpreadsheetML."""
    return '{'+S+'}'+name

def TestCanonical(node,preserve_text=False):
    """Сравнивает смысл XML без отличий префиксов и отступов сериализации."""
    text=(node.text or '') if preserve_text else (node.text or '').strip()
    return (node.tag,tuple(sorted(node.attrib.items())),text,tuple(TestCanonical(c,preserve_text) for c in node))

def TestMergeStyles(old,source):
    """Импортирует стили новой справки, сохраняя все старые ID Results и пользовательских ячеек."""
    names=['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs']
    targets={name:old.find(TestTag(name)) for name in names}
    sources={name:source.find(TestTag(name)) for name in names}
    assert all(targets[n] is not None and sources[n] is not None for n in names)
    before={n:[TestCanonical(c) for c in targets[n]] for n in names}
    ids={}
    existing={e.get('formatCode'):int(e.get('numFmtId')) for e in targets['numFmts']}
    next_id=max([163]+list(existing.values()))+1
    for e in sources['numFmts']:
        code=e.get('formatCode');number=int(e.get('numFmtId'))
        if code not in existing:
            c=deepcopy(e);c.set('numFmtId',str(next_id));targets['numFmts'].append(c);existing[code]=next_id;next_id+=1
        ids[number]=existing[code]
    def TestUnion(name,transform=None):
        """Добавляет отсутствующие записи, не меняя существующие; возвращает sourceID→readyID."""
        records=targets[name];index={TestCanonical(c):i for i,c in enumerate(records)};mapping={}
        for i,item in enumerate(sources[name]):
            clone=deepcopy(item)
            if transform:transform(clone)
            key=TestCanonical(clone)
            if key not in index:index[key]=len(records);records.append(clone)
            mapping[i]=index[key]
        return mapping
    fonts=TestUnion('fonts');fills=TestUnion('fills');borders=TestUnion('borders')
    def TestRemapXF(node,base=None):
        """Перепривязывает ссылки нового XF к добавленным старой книге коллекциям."""
        for name,mapping in [('fontId',fonts),('fillId',fills),('borderId',borders)]:
            if node.get(name) is not None:node.set(name,str(mapping[int(node.get(name))]))
        if node.get('numFmtId') is not None:
            value=int(node.get('numFmtId'));node.set('numFmtId',str(ids[value] if value>=164 else value))
        if base is not None and node.get('xfId') is not None:node.set('xfId',str(base[int(node.get('xfId'))]))
    bases=TestUnion('cellStyleXfs',TestRemapXF)
    cells=TestUnion('cellXfs',lambda node:TestRemapXF(node,bases))
    for name in names:
        targets[name].set('count',str(len(targets[name])))
        assert [TestCanonical(c) for c in targets[name]][:len(before[name])]==before[name],name
    return cells,{'PreservedStylePrefixes':{n:len(before[n]) for n in names},'FinalStyleCounts':{n:len(targets[n]) for n in names}}

def TestMapSheetStyles(sheet,mapping):
    """Привязывает новые ячейки, строки и столбцы к объединённым стилям Ready30."""
    for node in sheet.iter():
        attr='style' if node.tag==TestTag('col') else 's' if node.tag in (TestTag('c'),TestTag('row')) else None
        if attr and node.get(attr) is not None:node.set(attr,str(mapping[int(node.get(attr))]))
    assert not sheet.findall('s:conditionalFormatting',NS),'New static Help should not depend on unmerged conditional formats'

def TestInlineStrings(sheet,strings):
    """Встраивает новые тексты без перенумерации исходного sharedStrings и Results."""
    for cell in sheet.findall('.//s:sheetData/s:row/s:c',NS):
        if cell.get('t')!='s':continue
        assert cell.find(TestTag('f')) is None,'A shared string with a formula needs separate review'
        value=cell.find(TestTag('v'));record=strings[int(value.text)];cell.remove(value);cell.set('t','inlineStr')
        inline=ET.SubElement(cell,TestTag('is'))
        for child in record:inline.append(deepcopy(child))

def TestCells(sheet):
    """Индексирует реальную сетку по адресам; повтор адреса считается повреждением."""
    cells=sheet.findall('.//s:sheetData/s:row/s:c',NS);result={c.get('r'):c for c in cells};assert len(result)==len(cells)
    return result

def TestValue(cell,strings):
    """Читает точный текст/кеш ячейки без пересчёта и округления чисел."""
    if cell is None:return ''
    if cell.get('t')=='inlineStr':return ''.join(cell.find(TestTag('is')).itertext())
    v=cell.find(TestTag('v'))
    if v is None:return ''
    return ''.join(strings[int(v.text)].itertext()) if cell.get('t')=='s' else (v.text or '')

def TestPosition(address):
    """Разбирает один адрес SpreadsheetML, допускает абсолютные признаки."""
    m=re.fullmatch(r'\$?([A-Z]+)\$?(\d+)',address);assert m,address;column=0
    for ch in m[1]:column=column*26+ord(ch)-64
    return column,int(m[2])

def TestValidationCells(sqref):
    """Перечисляет небольшую реальную область списка; огромную неоднозначную область отклоняет."""
    result=set()
    for area in sqref.split():
        ends=area.split(':');a=TestPosition(ends[0]);b=TestPosition(ends[-1]);assert 0<=b[1]-a[1]<1000 and 0<=b[0]-a[0]<50
        for column in range(a[0],b[0]+1):
            number=column;letters=''
            while number:number,mod=divmod(number-1,26);letters=chr(65+mod)+letters
            result.update(letters+str(r) for r in range(a[1],b[1]+1))
    return result

def TestPayload(cell):
    """Сравнивает всю запись пользовательской ячейки, кроме её адреса при переносе комментария."""
    clone=deepcopy(cell);clone.attrib.pop('r',None);return TestCanonical(clone,True)

def TestReplaceCell(sheet,address,cell):
    """Заменяет запись без смещения строк/столбцов и сохраняет сортировку ячеек."""
    row_number=TestPosition(address)[1];rows=sheet.find(TestTag('sheetData'));row=next((r for r in rows if int(r.get('r'))==row_number),None)
    assert row is not None,'Expected existing Main row'
    for old in list(row):
        if old.get('r')==address:row.remove(old)
    cell.set('r',address);row.append(cell)
    ordered=sorted(list(row),key=lambda c:TestPosition(c.get('r'))[0]);row[:]=ordered

def TestPatchMain(old,new,old_strings,new_strings,style_map,definitions,section_bounds=None):
    """Добавляет четыре режима в H, переносит комментарии H→I и сохраняет все прочие данные Сечений."""
    old_cells=TestCells(old);new_cells=TestCells(new);before={a:TestPayload(c) for a,c in old_cells.items()};allowed=set();editable=set();moved=[];modes=[]
    for definition in definitions:
        block=next(b for b in definition['Blocks'] if b['Key']=='Main')
        keys=[str(r[0]) for r in block['Rows']]
        bounds=None if section_bounds is None else section_bounds[definition['RangeName']]
        candidates=[]
        for a,c in new_cells.items():
            if not a.startswith('H') or TestValue(c,new_strings)!='endMode':continue
            row=TestPosition(a)[1]
            if bounds is not None and not (bounds[0]<=row-1 and row+1+len(keys)<=bounds[1]):continue
            if [TestValue(new_cells.get('A'+str(row+2+i)),new_strings) for i in range(len(keys))]==keys:candidates.append(row)
        assert len(candidates)==1,(definition['Type'],candidates)
        machine=candidates[0]
        assert TestValue(old_cells.get('H'+str(machine)),old_strings)=='note'
        for row in [machine-1,machine,machine+1]+list(range(machine+2,machine+2+len(keys))):
            h='H'+str(row);i='I'+str(row);original=old_cells.get(h);assert original is not None
            assert TestValue(old_cells.get(i),old_strings)=='','Existing reserved I cell is not empty'
            assert old_cells.get(i) is None or old_cells[i].find(TestTag('f')) is None,'Reserved I contains a user formula'
            note=deepcopy(original);TestReplaceCell(old,i,note);moved.append((h,i,TestPayload(original)))
            mode=deepcopy(new_cells[h]);tiny=ET.Element(TestTag('worksheet'));data=ET.SubElement(tiny,TestTag('sheetData'));r=ET.SubElement(data,TestTag('row'));r.append(mode)
            TestInlineStrings(tiny,new_strings)
            if mode.get('s') is not None:mode.set('s',str(style_map[int(mode.get('s'))]))
            TestReplaceCell(old,h,mode);allowed.update((h,i))
        for j,key in enumerate(keys):
            address='H'+str(machine+2+j);value=TestValue(new_cells[address],new_strings)
            if key.startswith('Opening'):assert value=='—'
            else:
                assert value in ('IncludeBoth','ExcludeBoth','IncludeStart','IncludeEnd');editable.add(address)
            modes.append({'Type':definition['Type'],'Key':key,'Address':address,'Value':value})
    previous=old.find(TestTag('dataValidations'));current=new.find(TestTag('dataValidations'));assert previous is not None and current is not None
    for validation in previous:
        assert not (TestValidationCells(validation.get('sqref'))&allowed),'Old reserved Main note has a validation; review instead of discarding it'
    imported=set()
    for validation in current:
        cells=TestValidationCells(validation.get('sqref'))
        if not cells&editable:continue
        assert cells<=editable and not cells&imported
        assert validation.get('type')=='list'
        assert validation.findtext('s:formula1',namespaces=NS)=='"IncludeBoth,ExcludeBoth,IncludeStart,IncludeEnd"'
        previous.append(deepcopy(validation));imported.update(cells)
    assert imported==editable;previous.set('count',str(len(previous)))
    after=TestCells(old)
    for a,payload in before.items():
        if a not in allowed:assert TestPayload(after[a])==payload,a
    for h,i,payload in moved:assert TestPayload(after[i])==payload,(h,i)
    return {'ModeCells':modes,'PreservedOtherGeometryCells':len(before)-len(set(before)&allowed),'MovedCommentCells':len(moved),'FixedListModeCells':len(editable),'TableRowsAndColumnsInserted':False,'NamedSectionBounds':section_bounds}

def TestUpdateHelpLinks(old,new):
    """Обновляет только адреса существующих ссылок, не трогая настройки/формулы Config."""
    target={x.get('ref'):x.get('location') for x in new.findall('s:hyperlinks/s:hyperlink',NS) if (x.get('location') or '').startswith("'Справка'!")}
    links=[x for x in old.findall('s:hyperlinks/s:hyperlink',NS) if (x.get('location') or '').startswith("'Справка'!")]
    assert {x.get('ref') for x in links}==set(target),'Help links changed shape; review required'
    changes=[]
    for link in links:
        address=target[link.get('ref')]
        if link.get('location')!=address:changes.append({'Cell':link.get('ref'),'Before':link.get('location'),'After':address});link.set('location',address)
    return changes

def TestSelfTest():
    """Проверяет конфликтующие ID стилей, сохранность старых записей и точный Unicode в новых inline-текстах."""
    def TestStyle(font,code):
        return ET.fromstring(('<styleSheet xmlns="'+S+'"><numFmts count="1"><numFmt numFmtId="164" formatCode="'+code+'"/></numFmts><fonts count="1"><font><name val="'+font+'"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf fontId="0" fillId="0" borderId="0" numFmtId="0"/></cellStyleXfs><cellXfs count="1"><xf fontId="0" fillId="0" borderId="0" numFmtId="164" xfId="0"/></cellXfs></styleSheet>').encode())
    old=TestStyle('Arial','0.00');source=TestStyle('Calibri','0.000');original=TestCanonical(old.find('s:cellXfs/s:xf',NS));mapping,proof=TestMergeStyles(old,source)
    assert mapping=={0:1} and TestCanonical(old.find('s:cellXfs/s:xf',NS))==original
    imported=old.find('s:cellXfs',NS)[mapping[0]];assert imported.get('fontId')=='1' and imported.get('numFmtId')=='165'
    strings=ET.fromstring(('<sst xmlns="'+S+'"><si><t xml:space="preserve">  Отступы s₀; &lt;&amp;&gt;\nстрока 2  </t></si></sst>').encode())
    sheet=ET.fromstring(('<worksheet xmlns="'+S+'"><sheetData><row r="1"><c r="H1" s="0" t="s"><v>0</v></c></row></sheetData></worksheet>').encode())
    TestInlineStrings(sheet,strings);TestMapSheetStyles(sheet,mapping);assert TestValue(TestCells(sheet)['H1'],[])=='  Отступы s₀; <&>\nстрока 2  '
    assert TestCells(sheet)['H1'].get('s')=='1'
    assert TestValidationCells('H12 H36:H43')=={'H12'}|{'H'+str(n) for n in range(36,44)}
    assert TestValidationCells('G2:H3')=={'G2','G3','H2','H3'}
    left=ET.fromstring(('<c xmlns="'+S+'" r="A1" t="inlineStr"><is><t xml:space="preserve"> x </t></is></c>').encode())
    right=ET.fromstring(('<c xmlns="'+S+'" r="I1" t="inlineStr"><is><t>x</t></is></c>').encode())
    assert TestPayload(left)!=TestPayload(right),'Whitespace changes must be rejected'
    return {'Status':'SYNTHETIC_SELFTEST_PASS','RealWorkbookRead':False,'ExcelLaunched':False,'OldStylePrefixPreserved':True,'ConflictingStyleIDsRemapped':True,'UnicodeAndWhitespaceExact':True}

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--selftest',action='store_true');args=parser.parse_args()
    if not args.selftest:raise SystemExit('Pure helper only: use the guarded end-only preparation script after full PASS.')
    print(json.dumps(TestSelfTest(),ensure_ascii=False))
