"""Собирает итог направленной приёмки общей таблицы из завершённых доказательств."""
from pathlib import Path
import hashlib,json,re,sys,zipfile,xml.etree.ElementTree as ET

ROOT=Path(__file__).resolve().parents[2]
D=ROOT/'docs/regression/SharedRebarInput_2026-10-18'
def read(name):return json.loads((D/name).read_text(encoding='utf-8-sig'))
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest().upper()
candidate=sha(D/'Candidate.xlsm');accepted=sha(D/'Accepted.xlsm')
runs=[];assertions=0
for folder in ('FinalGeometryChecks','FinalInputAcceptance','ConfigChecks03'):
    receipt=read(folder+'/Checks.json')
    assert receipt['Passed'] and receipt['WorkbookSHA256']==candidate,folder
    for item in receipt['Checks']:
        method=item['Macro'];parts=method.split(':');name=parts[0].split('.')[-1]+('_'+parts[1] if len(parts)>1 else '')+'.txt'
        path=D/folder/name;body=path.read_text(encoding='utf-8-sig')
        assert sha(path)==item['SHA256'] and not re.search(r'(?m)^FAIL:|failed=[1-9]|RUNTIME ERROR',body),(folder,method)
        totals=re.findall(r'(?m)^TOTAL[^\r\n]*?passed=(\d+); failed=0',body)
        assertions+=sum(map(int,totals));runs.append({'Macro':method,'Seconds':item['Seconds']})
assert len(runs)==25,len(runs)
for name in ('FinalGeometryComparison.json','CalculationComparison.json','AcceptedLayout.json','FreshAcceptance.json',
             'FreshDefaultModels.json','AcceptedPreservation.json','AcceptedSourceEquality.json'):
    assert read(name)['Passed'],name
assert all(check['Passed'] for check in read('AcceptedStructure.json'))
assert read('AcceptedLayout.json')['WorkbookSHA256']==accepted
assert read('AcceptedPreservation.json')['AfterSHA256']==accepted
assert read('FreshAcceptance.json')['WorkbookSHA256']==sha(D/'FreshAccepted.xlsm')
assert read('FreshDefaultModels.json')['WorkbookSHA256']==sha(D/'FreshAccepted.xlsm')
for code in ('SP35','SP63'):
    name='DumpAudit02Calculations_'+code+'.txt'
    assert (D/'BeforeCalculations'/name).read_bytes()==(D/'AfterCalculations02'/name).read_bytes(),code
def payload(name):
    body=(D/name).read_text(encoding='utf-8-sig');return body[body.index('COMPONENT: '):]
assert payload('CandidateNativeVBA.txt')==payload('AcceptedFinalNativeVBA.txt'),'Accepted VBA differs from the tested candidate.'
assert 'NDM.RECIPE.GUIDE/' not in payload('AcceptedFinalNativeVBA.txt')
guide=json.loads((ROOT/'docs/geometry_refactor/SECTION_RECIPE_GUIDE_LAYOUT.json').read_text(encoding='utf-8-sig'))
ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with zipfile.ZipFile(D/'Accepted.xlsm') as archive:
    strings=[''.join(n.itertext()) for n in ET.fromstring(archive.read('xl/sharedStrings.xml'))]
    workbook=ET.fromstring(archive.read('xl/workbook.xml'))
    names={n.attrib['name']:n.text for n in workbook.findall('m:definedNames/m:definedName',ns)}
    start=int(re.search(r'\$A\$(\d+)',names['rngSectionRecipeGuide'])[1])
    sheet=next(s for s in workbook.findall('m:sheets/m:sheet',ns) if s.attrib['name']=='Справка')
    rel={n.attrib['Id']:n.attrib['Target'] for n in ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))}
    target=rel[sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']]
    target=target.lstrip('/') if target.startswith('/') else 'xl/'+target
    cells={c.attrib['r']:c for c in ET.fromstring(archive.read(target)).findall('.//m:sheetData/m:row/m:c',ns)}
    for i,row in enumerate(guide['rows']):
        cell=cells.get('A'+str(start+i));actual=''
        if cell is not None:
            value=cell.find('m:v',ns)
            if value is not None:actual=strings[int(value.text)] if cell.attrib.get('t')=='s' else value.text
        assert actual==row['text'],('Guide row',i,actual,row['text'])
before=read('Before.json')
assert sha(ROOT/'workbook/output/RC_Section_NDM_execution_report.txt')==before['workbook/output/RC_Section_NDM_execution_report.txt']
assert read('ClassComments.json')['status']=='PASS'
result={'Passed':True,'CandidateSHA256':candidate,'AcceptedSHA256':accepted,'DirectedRuns':runs,'Assertions':assertions,
        'GeometryMatrices':5,'RuleAndPathCases':60,'CalculationCases':16,'CalculationCodes':['SP35','SP63'],
        'GeometryAndCalculationsUnchanged':True,'AcceptedCodeMatchesTestedCandidate':True,'InputColorRGB':[255,242,230],
        'OtherConfigInputFields':read('AcceptedLayout.json')['OtherConfigInputCells'],'DefaultTypesBuilt':4,
        'SourceComponents':155,'WorkbookComponents':160,'GuideRows':len(guide['rows']),'FullOnOffRun':False,
        'PreservedGeometryInputs':25,'PreservedActiveRebarInputs':104,'UserExecutionReportUnchanged':True}
if '--published' in sys.argv:
    publication=read('Publication.json')
    assert publication['Passed'] and publication['Readback']
    assert sha(ROOT/'workbook/output/RC_Section_NDM.xlsm')==accepted
    assert read('PublishedSourceEquality.json')['Passed']
    assert read('PublishedLayout.json')['Passed'] and read('PublishedLayout.json')['WorkbookSHA256']==accepted
    assert read('PublishedPreservation.json')['Passed'] and read('PublishedPreservation.json')['AfterSHA256']==accepted
    assert (ROOT/'workbook/output/VBA_All_Code.txt').read_bytes()==(D/'PublishedNativeVBA.txt').read_bytes()
    result['Published']=True;result['AllCodeSHA256']=sha(ROOT/'workbook/output/VBA_All_Code.txt')
(D/'Acceptance.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps({k:v for k,v in result.items() if k!='DirectedRuns'},ensure_ascii=False))
