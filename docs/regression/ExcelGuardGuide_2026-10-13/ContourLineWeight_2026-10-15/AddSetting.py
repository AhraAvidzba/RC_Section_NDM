"""Точечно добавляет поле и его справку, сохраняя все прежние данные/оформление ZIP."""
import copy
import json
import re
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as ET

directory=Path(__file__).resolve().parent
S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
with ZipFile(directory/'Baseline.xlsm') as baseline, ZipFile(directory/'Candidate.xlsm') as candidate:
    parts={name:baseline.read(name) for name in baseline.namelist()}
    parts['xl/vbaProject.bin']=candidate.read('xl/vbaProject.bin')
    strings=[''.join(item.itertext()) for item in ET.fromstring(parts['xl/sharedStrings.xml'])]
    def value(cell):
        if cell is None:return None
        v=cell.find(S+'v')
        if cell.get('t')=='s':return strings[int(v.text)] if v is not None else None
        if cell.get('t')=='inlineStr':return ''.join(cell.find(S+'is').itertext())
        return v.text if v is not None else None
    def assign(cell,text):
        for child in list(cell):
            if child.tag in (S+'f',S+'v',S+'is'):cell.remove(child)
        cell.attrib.pop('t',None)
        if text is None:return
        if isinstance(text,(int,float)):ET.SubElement(cell,S+'v').text=str(text)
        else:
            cell.set('t','inlineStr');ET.SubElement(ET.SubElement(cell,S+'is'),S+'t').text=text
    book=ET.fromstring(parts['xl/workbook.xml'])
    names={e.get('name'):e for e in book.find(S+'definedNames')}
    relationships={e.get('Id'):e.get('Target') for e in ET.fromstring(parts['xl/_rels/workbook.xml.rels'])}
    def sheet(name):
        item=next(e for e in book.find(S+'sheets') if e.get('name')==name)
        target=relationships[item.get(R+'id')]
        target=target.lstrip('/') if target.startswith('/') else 'xl/'+target
        return target,ET.fromstring(parts[target])
    config_part,config=sheet('Config');help_part,help_sheet=sheet('Справка')
    caption_changes=[]
    for sheet_name,document in [('Config',config),('Справка',help_sheet)]:
        for cell in document.iter(S+'c'):
            text=value(cell)
            if isinstance(text,str) and 'Аннотации схемы' in text:
                new=text.replace('Аннотации схемы','Настройки оформления схемы')
                new=new.replace('на Config управляет размерными линиями и подписями арматуры на Excel-схеме.','на Config управляет размерами, подписями арматуры и толщиной границ сечения на Excel-схеме.')
                assign(cell,new);caption_changes.append(sheet_name+'!'+cell.get('r'))
    cells={e.get('r'):e for e in config.iter(S+'c')}
    header,end=map(int,re.findall(r'\$(\d+)',names['rngPlotAnnotationSettings'].text))
    assert end-header==13
    row=end+1
    assert all(value(cells.get(col+str(row))) is None for col in 'HIJKLM')
    source_row=next(int(re.search(r'\d+',address)[0]) for address,cell in cells.items() if address.startswith('H') and value(cell)=='LineWeight' and header<int(re.search(r'\d+',address)[0])<=end)
    rows=config.find(S+'sheetData')
    target_row=next(e for e in rows if int(e.get('r'))==row)
    comment='Общая толщина наружного контура и границ отверстий на Excel-схеме. Прямые, дуги и окружности используют одно значение > 0; по умолчанию 2 pt.'
    def col_number(address):
        result=0
        for char in re.match('[A-Z]+',address)[0]:result=result*26+ord(char)-64
        return result
    def put_row(document,number,prototypes,values,height=None):
        data=document.find(S+'sheetData')
        current=next((e for e in data if int(e.get('r'))==number),None)
        if current is None:
            current=ET.Element(S+'row',r=str(number));data.append(current)
        if height:
            current.set('ht',str(height));current.set('customHeight','1')
        for col,prototype in prototypes.items():
            address=col+str(number)
            old=next((e for e in current if e.get('r')==address),None)
            if old is not None:current.remove(old)
            cell=copy.deepcopy(prototype);cell.set('r',address);assign(cell,values.get(col));current.append(cell)
        current[:]=sorted(current,key=lambda e:col_number(e.get('r')))
        data[:]=sorted(data,key=lambda e:int(e.get('r')))
    put_row(config,row,{col:cells[col+str(source_row)] for col in 'HIJKLM'},dict(zip('HIJKLM',['ContourLineWeight',2,None,'pt',comment,'Подробнее'])))
    merge=config.find(S+'mergeCells');ET.SubElement(merge,S+'mergeCell',ref=f'I{row}:J{row}');merge.set('count',str(len(merge)))
    validation=config.find(S+'dataValidations')
    item=ET.SubElement(validation,S+'dataValidation',type='decimal',operator='greaterThan',allowBlank='0',showErrorMessage='1',showInputMessage='1',sqref=f'I{row}:J{row}',errorTitle='Толщина контура',error='Введите положительное число в пунктах (pt).')
    ET.SubElement(item,S+'formula1').text='0';validation.set('count',str(len(validation)))
    names['rngPlotAnnotationSettings'].text=re.sub(r'\$\d+$',lambda _: '$'+str(row),names['rngPlotAnnotationSettings'].text)

    help_cells={e.get('r'):e for e in help_sheet.iter(S+'c')}
    title=next(address for address,cell in help_cells.items() if value(cell)=='Настройки оформления схемы: LineWeight')
    prototype=int(re.search(r'\d+',title)[0])
    last=max(int(re.search(r'\d+',address)[0]) for address,cell in help_cells.items() if value(cell) is not None)
    start=last+2
    assert all(value(help_cells.get(col+str(n))) is None for n in range(start,start+3) for col in 'ABCDEF')
    put_row(help_sheet,start,{col:help_cells[col+str(prototype)] for col in 'ABCDEF'},{'A':'Настройки оформления схемы: ContourLineWeight'},24)
    lines=[
        'ContourLineWeight — общая толщина наружных границ и границ отверстий на Excel-схеме, в пунктах (pt). Значение по умолчанию — 2 pt.',
        'Одинаково применяется к прямым, дугам и окружностям у созданных программой и импортированных сечений. Не зависит от единиц длины и масштаба диаграммы. Введите положительное число; при Plot.ContourEnabled = No значение не используется.',
    ]
    for i,text in enumerate(lines,1):
        put_row(help_sheet,start+i,{col:help_cells[col+str(prototype+1)] for col in 'ABCDEF'},{'A':str(i),'B':text},31.5)
    merge=help_sheet.find(S+'mergeCells')
    for first,last_col,number in [('A','F',start),('B','F',start+1),('B','F',start+2)]:
        ET.SubElement(merge,S+'mergeCell',ref=f'{first}{number}:{last_col}{number}')
    merge.set('count',str(len(merge)))
    ET.SubElement(config.find(S+'hyperlinks'),S+'hyperlink',ref=f'M{row}',location=f"'Справка'!A{start}",display='Подробнее')
    help_links=help_sheet.find(S+'hyperlinks')
    if help_links is None:
        help_links=ET.Element(S+'hyperlinks')
        margins=help_sheet.find(S+'pageMargins')
        assert margins is not None
        help_sheet.insert(help_sheet.index(margins),help_links)
    ET.SubElement(help_links,S+'hyperlink',ref=f'A{start}',location=f"'Config'!I{row}",display='Настройки оформления схемы: ContourLineWeight')
    parts[config_part]=ET.tostring(config,encoding='utf-8',xml_declaration=True)
    parts[help_part]=ET.tostring(help_sheet,encoding='utf-8',xml_declaration=True)
    parts['xl/workbook.xml']=ET.tostring(book,encoding='utf-8',xml_declaration=True)
    temp=directory/'Candidate.patched.tmp'
    with ZipFile(temp,'w',ZIP_DEFLATED) as out:
        for entry in baseline.infolist():out.writestr(entry,parts[entry.filename])
temp.replace(directory/'Candidate.xlsm')
result={'Passed':True,'Default':2,'Unit':'pt','InputCell':f'Config!I{row}','InputMerge':f'I{row}:J{row}',
        'TableName':'rngPlotAnnotationSettings','TableRange':names['rngPlotAnnotationSettings'].text,'HelpStartRow':start,
        'NewConfigCells':[col+str(row) for col in 'HIJKLM'],'NewHelpRows':list(range(start,start+3)),
        'CaptionChanges':caption_changes,'PreservedOtherZipParts':True}
(directory/'SettingAdded.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(result,ensure_ascii=False))
