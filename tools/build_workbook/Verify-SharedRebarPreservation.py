"""Сверяет сохранность данных вне заменённых таблиц и перенос исходных значений."""
from pathlib import Path
from decimal import Decimal
import argparse,hashlib,json,re,zipfile,xml.etree.ElementTree as ET

NS={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
def column(s):
    n=0
    for c in s:n=n*26+ord(c)-64
    return n
def address(s):
    m=re.fullmatch(r'\$?([A-Z]+)\$?(\d+)',s)
    return int(m[2]),column(m[1])
class Book:
    def __init__(self,path):
        with zipfile.ZipFile(path) as z:
            strings=[]
            if 'xl/sharedStrings.xml' in z.namelist():
                strings=[''.join(n.itertext()) for n in ET.fromstring(z.read('xl/sharedStrings.xml'))]
            w=ET.fromstring(z.read('xl/workbook.xml'))
            rel={n.attrib['Id']:n.attrib['Target'] for n in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
            self.names={n.attrib['name']:n.text for n in w.findall('m:definedNames/m:definedName',NS)}
            self.sheets={}
            for sheet in w.findall('m:sheets/m:sheet',NS):
                target=rel[sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']]
                path=target.lstrip('/') if target.startswith('/') else 'xl/'+target
                cells={}
                for c in ET.fromstring(z.read(path)).findall('.//m:sheetData/m:row/m:c',NS):
                    f=c.find('m:f',NS);v=c.find('m:v',NS);kind=c.attrib.get('t','n')
                    if f is not None:value=('formula',f.text)
                    elif kind=='s':value=('text',strings[int(v.text)])
                    elif kind=='inlineStr':value=('text',''.join(c.find('m:is',NS).itertext()))
                    elif v is not None and v.text is not None:value=(kind,Decimal(v.text) if kind=='n' else v.text)
                    else:continue
                    if value[1] not in ('',None):cells[address(c.attrib['r'])]=value
                self.sheets[sheet.attrib['name']]=cells
    def bounds(self,name):
        ref=self.names[name];sheet,area=ref.rsplit('!',1);start,end=area.split(':')
        return sheet.strip("'"),address(start),address(end)
    def value(self,sheet,row,col):return self.sheets[sheet].get((row,col),('empty',None))[1]

p=argparse.ArgumentParser();p.add_argument('--before',required=True);p.add_argument('--after',required=True);p.add_argument('--inputs',required=True);p.add_argument('--report',required=True)
a=p.parse_args();old,new=Book(a.before),Book(a.after)
assert old.sheets.keys()==new.sheets.keys(),'Worksheet list changed.'
unchanged=[]
for name in old.sheets:
    if name in ('Config','Справка'):continue
    assert old.sheets[name]==new.sheets[name],f'Unrelated sheet changed: {name}'
    unchanged.append(name)
geometry=['rngCircleGeometry','rngRectSetGeometry','rngPointedRectGeometry','rngHollowRectangleGeometry']
preserved_tables=['rngSystemSettings','rngUnitSettings','rngSignConventionSettings','rngSteelMaterialParameters',
    'rngConcreteMaterialParameters','rngCalculationProfiles','rngPlotAnnotationSettings','rngLoadCombinations',
    'rngStabilityDurationLoads','rngSP35Table721']
for name in preserved_tables:
    assert old.bounds(name)==new.bounds(name),(name,'moved')
    sheet,(r1,c1),(r2,c2)=old.bounds(name)
    for row in range(r1,r2+1):
        for col in range(c1,c2+1):assert old.value(sheet,row,col)==new.value(sheet,row,col),(name,row,col)
inputs=json.loads(Path(a.inputs).read_text(encoding='utf-8-sig'))
def blocks(book,name):
    sheet,(first,col),(last,right)=book.bounds(name);result={}
    keyrows=[r for r in range(first,last+1) if book.value(sheet,r,col)=='key']
    for i,keyrow in enumerate(keyrows):
        block=book.value(sheet,keyrow-2,col)
        fields={book.value(sheet,keyrow,c):c for c in range(col,right+1) if book.value(sheet,keyrow,c)}
        result[block]={}
        end=keyrows[i+1]-4 if i+1<len(keyrows) else last
        for row in range(keyrow+1,end+1):
            key=book.value(sheet,row,col)
            if key and key!='Units':result[block][key]={k:book.value(sheet,row,c) for k,c in fields.items()}
    return result
count=0
for name in geometry:
    for key,fields in blocks(new,name)['General'].items():
        source='Geom|'+key+'|val'
        if key.startswith(('Left.','Right.')):
            side,key=key.split('.');source='Detail|'+key+'|'+side.lower()
        elif source not in inputs[name]:source='Detail|'+key+'|val'
        assert fields['val']==inputs[name][source],(name,key,fields['val'],inputs[name][source])
        count+=1
system_name='rngSystemSettings';sheet,(start,col),(end,_)=new.bounds(system_name)
selected=next(new.value(sheet,r,col+1) for r in range(start,end+1) if new.value(sheet,r,col)=='Geometry.Type')
rebar_count=0
for block,rows in blocks(new,'rngRebarInput').items():
    # Formula keys are checked natively. Stored source positions identify the active rows here.
    sheet,(start,col),(end,right)=new.bounds('rngRebarInput')
    keyrows=[r for r in range(start,end+1) if new.value(sheet,r,col)=='key']
    keyrow=keyrows[0 if block=='Main' else 1]
    prior=blocks(old,'rng'+selected+'Geometry')[block]
    for i,key in enumerate(prior):
        for field in prior[key]:
            if field in ('key','name','note','unit'):continue
            column_index=next(c for c in range(col,right+1) if new.value(sheet,keyrow,c)==field)
            value=new.value(sheet,keyrow+2+i,column_index)
            expected=inputs['rng'+selected+'Geometry'][block+'|'+key+'|'+field]
            if expected in ('-','—'):expected=None
            assert value==expected,(block,key,field,value,expected)
            rebar_count+=1
result={'Passed':True,'BeforeSHA256':hashlib.sha256(Path(a.before).read_bytes()).hexdigest().upper(),
    'AfterSHA256':hashlib.sha256(Path(a.after).read_bytes()).hexdigest().upper(),
    'UnchangedSheets':unchanged,'OtherConfigTablesUnchanged':preserved_tables,'GeometryInputsPreserved':count,'SelectedType':selected,'RebarInputsPreserved':rebar_count}
Path(a.report).write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(result,ensure_ascii=False))
