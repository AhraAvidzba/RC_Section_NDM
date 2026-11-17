"""ДЛЯ ТЕСТОВ: пассивно сверяет геометрию и оформление листа Расчет с main."""
from pathlib import Path
from zipfile import ZipFile
import json, xml.etree.ElementTree as E
R=Path(r'C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM')
P=Path(r'C:/Users/avidzba/.codex/visualizations/2026/10/12/01a12ece-8192-7141-89b3-f83f1176184c/TorsionPostprocess_2026-11-17')
N={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def canon(e):
    if e is None: return None
    return (e.tag,tuple(sorted(e.attrib.items())),e.text or '',tuple(canon(c) for c in e))
def inspect(path):
    with ZipFile(path) as z:
        book=E.fromstring(z.read('xl/workbook.xml'))
        sheets=book.find('m:sheets',N)
        index=next(i for i,s in enumerate(sheets) if s.get('name')=='Расчет')
        rid=sheets[index].get('{'+N['r']+'}id')
        rel=E.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        target=next(x.get('Target') for x in rel if x.get('Id')==rid)
        part='xl/'+target if not target.startswith('/') else target.lstrip('/')
        sheet=E.fromstring(z.read(part))
        styles=E.fromstring(z.read('xl/styles.xml'))
        def style(i):
            xf=styles.find('m:cellXfs',N)[int(i)]
            attrs=dict(xf.attrib)
            links={}
            for key,table in [('fontId','fonts'),('fillId','fills'),('borderId','borders')]:
                links[key]=canon(styles.find('m:'+table,N)[int(attrs.pop(key,'0'))])
            num=attrs.pop('numFmtId','0')
            nf=styles.find('m:numFmts',N)
            found=next((x.get('formatCode') for x in (list(nf) if nf is not None else []) if x.get('numFmtId')==num),None)
            links['numFmtId']=found or num
            attrs.pop('xfId',None)
            return (tuple(sorted(attrs.items())),links,tuple(canon(x) for x in xf))
        rows=sheet.find('m:sheetData',N)
        cell_styles={c.get('r'):style(c.get('s','0')) for row in rows for c in row}
        row_sizes={row.get('r'):{k:v for k,v in row.attrib.items() if k in ('ht','hidden','customHeight','outlineLevel','collapsed')} for row in rows}
        print_names=[(x.get('name'),x.text) for x in book.find('m:definedNames',N) if x.get('localSheetId')==str(index) and 'Print_' in x.get('name','')]
        fields={key:canon(sheet.find('m:'+key,N)) for key in ['cols','sheetFormatPr','mergeCells','pageMargins','pageSetup','printOptions']}
        fields['mergeCells']=tuple(sorted(c.get('ref') for c in sheet.find('m:mergeCells',N)))
        fields['cols']=tuple((tuple(sorted((k,v) for k,v in col.attrib.items() if k!='style')),style(col.get('style','0'))) for col in sheet.find('m:cols',N))
        return dict(fields=fields,rows=row_sizes,styles=cell_styles,printNames=print_names)
a=inspect(R/'workbook/output/RC_Section_NDM.xlsm')
b=inspect(P/'FinalCandidate12/RC_Section_NDM.xlsm')
checks={k:a['fields'][k]==b['fields'][k] for k in a['fields']}
checks['PrintNames']=a['printNames']==b['printNames']
rows=set(a['rows'])|set(b['rows'])
rowdiff=[r for r in rows if a['rows'].get(r,{})!=b['rows'].get(r,{})]
cells=set(a['styles'])|set(b['styles'])
stylediff=[r for r in cells if a['styles'].get(r)!=b['styles'].get(r)]
result={'Checks':checks,'RowDifferences':rowdiff,'CellStyleDifferences':stylediff,'Equal':all(checks.values()) and not rowdiff and not stylediff}
(P/'CalcLayoutFinal12.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({k:v for k,v in result.items() if k not in ('CellStyleDifferences','RowDifferences')},ensure_ascii=False)); print('Rows=',len(rowdiff),'Styles=',len(stylediff))
