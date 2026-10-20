"""Сверяет пользовательские данные до и после переноса таблиц на лист «Сечения»."""
from __future__ import annotations
import argparse, hashlib, json, posixpath, re, zipfile
from decimal import Decimal
from pathlib import Path
from xml.etree import ElementTree as E
N={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

def read(path):
 with zipfile.ZipFile(path) as z:
  wb=E.fromstring(z.read('xl/workbook.xml'))
  rel={x.get('Id'):x.get('Target') for x in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
  ss=[''.join(x.itertext()) for x in E.fromstring(z.read('xl/sharedStrings.xml')).findall('s:si',N)] if 'xl/sharedStrings.xml' in z.namelist() else []
  names={x.get('name'):x.text for x in wb.findall('s:definedNames/s:definedName',N)}
  sheets={}
  for sheet in wb.find('s:sheets',N):
   node=E.fromstring(z.read(posixpath.normpath('xl/'+rel[sheet.get(R)])))
   cells={}
   for c in node.findall('s:sheetData/s:row/s:c',N):
    v=c.findtext('s:v',default='',namespaces=N);t=c.get('t');f=c.find('s:f',N)
    if t=='s':v=ss[int(v)]
    elif t=='inlineStr':v=''.join(c.find('s:is',N).itertext())
    elif v and t not in ('e','str','b'):v=str(Decimal(v).normalize())
    if f is not None:cells[c.get('r')]=('formula',f.text or '',tuple(sorted(f.attrib.items())))
    elif v!='':cells[c.get('r')]=(t if t in ('e','b') else 'value',v)
   sheets[sheet.get('name')]={'cells':cells,'cols':[x.attrib for x in node.findall('s:cols/s:col',N)]}
  calc=wb.find('s:calcPr',N)
  return names,sheets,calc.get('calcMode','auto')

def coord(ref):
 a,b=re.fullmatch(r'([A-Z]+)(\d+)',ref).groups();c=0
 for ch in a:c=c*26+ord(ch)-64
 return int(b),c

def bounds(ref):
 pts=re.findall(r'\$([A-Z]+)\$(\d+)',ref)
 return [coord(a+b) for a,b in (pts[0],pts[-1])]

def run(before,after):
 oldnames,old,oldcalc=read(before);newnames,new,newcalc=read(after)
 assert list(new)==['Config','Сечения','Справка','Расчет','Results']
 assert oldcalc==newcalc,(oldcalc,newcalc)
 excluded=[bounds(oldnames[k]) for k in ('rngCircleGeometry','rngRectanglesGeometry','rngRoundedRectangleGeometry','rngHollowRectangleGeometry','rngRebarInput')]
 # Включаем заголовок каждой переносимой таблицы и её исходные данные.
 excluded=[((a[0]-1,a[1]),b) for a,b in excluded]
 checked={}
 for name in ('Config','Расчет','Results'):
  count=0
  def columns(sheet):
   limit=69 if name=='Config' else 90
   return [{k:v for k,v in next((x for x in sheet['cols'] if int(x['min'])<=i<=int(x['max'])),{}).items() if k in ('width','hidden','customWidth')} for i in range(1,limit+1)]
  assert columns(old[name])==columns(new[name]),f'User widths changed on {name}'
  for ref in old[name]['cells'].keys() | new[name]['cells'].keys():
   r,c=coord(ref)
   if name=='Config' and (c>=70 or any(a[0]<=r<=b[0] and a[1]<=c<=b[1] for a,b in excluded)):continue
   a=old[name]['cells'].get(ref);b=new[name]['cells'].get(ref)
   if a!=b:
    assert a and b and a[0]==b[0]=='value' and a[1].replace('RoundedRectangle','PointedSection')==b[1],(name,ref,a,b)
   count+=1
  checked[name]=count
 removed=sorted(k for k in oldnames if k.endswith('_Choices'))
 assert len(removed)==35 and not any(k.endswith('_Choices') for k in newnames)
 assert 'rngRebarInput' not in newnames and 'rngRoundedRectangleGeometry' not in newnames
 assert all(k in newnames for k in ('rngCircleGeometry','rngRectanglesGeometry','rngPointedSectionGeometry','rngHollowRectangleGeometry'))
 # Все остальные именованные диапазоны сохраняют свои назначения.
 for key,value in oldnames.items():
  if key.startswith('_xlnm.') or key in removed or key in ('rngRebarInput','rngCircleGeometry','rngRectanglesGeometry','rngRoundedRectangleGeometry','rngHollowRectangleGeometry','rngSectionRecipeTypes'):continue
  assert newnames.get(key)==value,(key,value,newnames.get(key))
 return {'Status':'PASS','BeforeSHA256':hashlib.sha256(Path(before).read_bytes()).hexdigest().upper(),'AfterSHA256':hashlib.sha256(Path(after).read_bytes()).hexdigest().upper(),'UnchangedCells':checked,'UserColumnWidthsPreserved':True,'CalculationMode':newcalc,'RemovedChoiceNames':removed,'NamesBefore':len(oldnames),'NamesAfter':len(newnames),'OtherNamedRangesPreserved':True,'HiddenHelperCellsBefore':sum(coord(k)[1]>=70 for k in old['Config']['cells']),'HiddenHelperCellsAfter':sum(coord(k)[1]>=70 for k in new['Config']['cells'])}

if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--before',required=True);p.add_argument('--after',required=True);p.add_argument('--report',required=True);a=p.parse_args()
 result=run(a.before,a.after);Path(a.report).write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf8');print(json.dumps(result,ensure_ascii=False))
