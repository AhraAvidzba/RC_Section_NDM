"""Приёмка добавленного поля, названия блока и сохранности исходных данных книги."""
import argparse
import hashlib
import importlib.util
import json
import re
from pathlib import Path
from zipfile import ZipFile

directory=Path(__file__).resolve().parent
root=directory.parents[3]
spec=importlib.util.spec_from_file_location('presentation',root/'tools/build_workbook/Verify-GuidePresentation.py')
p=importlib.util.module_from_spec(spec);spec.loader.exec_module(p)
parser=argparse.ArgumentParser()
parser.add_argument('--book',type=Path,default=directory/'Candidate.xlsm')
parser.add_argument('--dump',type=Path,default=directory/'FinalNativeVBA.txt')
parser.add_argument('--report',type=Path,default=directory/'Preservation.json')
args=parser.parse_args()
old,new=p.reader.read(directory/'Baseline.xlsm'),p.reader.read(args.book)
edit=json.loads((directory/'SettingAdded.json').read_text(encoding='utf-8'))
errors=[]
def require(ok,text):
    if not ok:errors.append(text)
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
require(old['names'].keys()==new['names'].keys(),'Named-range inventory changed')
require(new['names']['rngPlotAnnotationSettings']==edit['TableRange'],'Annotation range not extended')
require(all(value==new['names'][key] for key,value in old['names'].items() if key!='rngPlotAnnotationSettings'),'Unrelated name changed')
approved=set(edit['CaptionChanges'])
approved|={'Config!'+address for address in edit['NewConfigCells']}
approved|={'Справка!'+col+str(row) for row in edit['NewHelpRows'] for col in 'ABCDEF'}
checked=0
require(old['sheets'].keys()==new['sheets'].keys(),'Sheet inventory changed')
for name,sheet in old['sheets'].items():
    current=new['sheets'][name]
    for address in sheet['cells'].keys()|current['cells'].keys():
        checked+=1
        a,b=sheet['cells'].get(address),current['cells'].get(address)
        if name+'!'+address in approved:
            if name+'!'+address in edit['CaptionChanges']:
                expected=a['value'].replace('Аннотации схемы','Настройки оформления схемы').replace('на Config управляет размерными линиями и подписями арматуры на Excel-схеме.','на Config управляет размерами, подписями арматуры и толщиной границ сечения на Excel-схеме.')
                require(b['value']==expected,'Wrong caption: '+address)
                require(all(a[k]==b[k] for k in ('formula','style','attrs')),'Caption formatting changed: '+address)
        else:require(a==b,'Existing cell changed: '+name+'!'+address)
    for row in sheet['rows'].keys()|current['rows'].keys():
        a,b=dict(sheet['rows'].get(row,{})),dict(current['rows'].get(row,{}))
        if name=='Справка' and row in edit['NewHelpRows']:
            require(0<float(b.get('ht',0))<=32,'New help row too tall')
            for key in ('ht','customHeight'):a.pop(key,None);b.pop(key,None)
        require(a==b,'Row metadata changed: '+name+'!'+str(row))
require(old['drawings']==new['drawings'],'Saved plot or another drawing changed')
config=new['sheets']['Config']['cells']
require(config['I84']['value']=='2' and config['K84']['value']=='pt','Actual default/unit is not 2 pt')
require(config['H69']['value']=='Настройки оформления схемы (Подробнее)','Wrong table title')
with ZipFile(directory/'Baseline.xlsm') as a,ZipFile(args.book) as b:
    require(set(a.namelist())==set(b.namelist()),'ZIP part inventory changed')
    changed=[name for name in a.namelist() if a.read(name)!=b.read(name)]
    require(set(changed)=={'xl/workbook.xml','xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml','xl/vbaProject.bin'},'Unrelated ZIP changes: '+str(changed))
    xml=p.reader.ET.fromstring(b.read('xl/worksheets/sheet1.xml'))
    S=p.reader.S
    require(any(e.get('ref')==edit['InputMerge'] for e in xml.iter(S+'mergeCell')),'Shared input not merged')
    require(any(e.get('sqref')==edit['InputMerge'] and e.get('type')=='decimal' and e.get('operator')=='greaterThan' and e.find(S+'formula1').text=='0' for e in xml.iter(S+'dataValidation')),'Missing positive decimal validation')
    require(any(e.get('ref')=='M84' and e.get('location')=="'Справка'!A1763" for e in xml.iter(S+'hyperlink')),'Setting link does not lead to its help')
    xml=p.reader.ET.fromstring(b.read('xl/worksheets/sheet2.xml'))
    require(any(e.get('ref')=='A1763' and e.get('location')=="'Config'!I84" for e in xml.iter(S+'hyperlink')),'Help return link missing')
baseline,final=p.components(directory/'BaselineNativeVBA.txt'),p.components(args.dump)
require(baseline.keys()==final.keys() and len(final)==157,'Native component inventory changed')
def normalize(text):
    # VBE меняет регистр имён при импорте; литералы и комментарии сохраняются.
    return ''.join(part if part.startswith(('"',"'")) else part.lower() for part in re.findall(r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27\n]+|\n',text.replace('\r\n','\n')))
changed_native=[name for name in baseline if normalize(baseline[name][2])!=normalize(final[name][2])]
require(set(changed_native)=={'CSectionPlotter','CSystemSettingsReader','modTestPlotConfig'},'Native code outside scope: '+str(changed_native))
result={'Passed':not errors,'CellsChecked':checked,'NamedRanges':len(new['names']),
        'ExistingValuesFormulasStylesAndDrawingsPreserved':not errors,'ChangedZIPParts':changed,
        'ChangedNativeComponents':changed_native,'Default':2,'Unit':'pt','NewTitle':'Настройки оформления схемы',
        'WorkbookSHA256':sha(args.book),'NativeSHA256':sha(args.dump),'Errors':errors}
args.report.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(result,ensure_ascii=False))
if errors:raise SystemExit(1)
