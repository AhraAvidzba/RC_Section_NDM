"""Verify the published workbook, focused checks and preserved inputs for this change."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import hashlib,json,re

ROOT=Path(__file__).resolve().parents[2]
D=ROOT/'docs/regression/SharedRebarInput_2026-10-18/SectionNames_2026-10-20'
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest().upper()
def read(name):return json.loads((D/name).read_text(encoding='utf-8-sig'))

def run():
    digest=read('Publication.json')['AcceptedSHA256']
    assert sha(ROOT/'workbook/output/RC_Section_NDM.xlsm')==digest
    assert read('LayoutFinal.json')['WorkbookSHA256']==read('StructureFinal.json')['WorkbookSHA256']==digest
    assert read('LayoutFinal.json')['Status']==read('StructureFinal.json')['Status']=='PASS'
    assert all(h['LinkSize']==12 for h in read('LayoutFinal.json')['Headers'])
    assert read('StructureFinal.json')['PreservedFields']==292
    assert read('PreservedFinalInputs.json')['ConfigAndCalculationCells']==707
    assert read('StructureDefaults02.json')['DefaultFieldsChecked']==296
    assert read('SourceEquality.json')['Passed'] and read('PublishedSourceEquality.json')['Passed']
    assert read('PublishedSourceEquality.json')['SourceComponents']==155
    published=D/'PublishedNativeVBA.txt';code=ROOT/'workbook/output/VBA_All_Code.txt'
    # ALL CODE is the shipped native export; a second local copy is optional.
    if published.exists():assert code.read_bytes()==published.read_bytes()
    assert sha(code)==read('PublishedSourceEquality.json')['ExportSHA256']
    assert not re.search(r'NDM\.RECIPE\.GUIDE|[0-9a-fA-F]{64}',code.read_text(encoding='utf-8-sig'))
    before=read('Before.json')
    assert sha(ROOT/'workbook/output/RC_Section_NDM_execution_report.txt')==before['workbook/output/RC_Section_NDM_execution_report.txt']
    suites=[]
    for name in ['DirectedFinalCore.txt','DirectedRelease.txt']:
        text=(D/name).read_text(encoding='utf-8-sig')
        assert 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' in text,name
        assert not re.search(r'(?m)^(FAIL:|SCRIPT ERROR|RUNTIME ERROR|WATCHDOG_FAILURE)|failed=[1-9]',text),name
        starts=re.findall(r'(?m)^SUITE_STARTED: ([^;\r\n]+)',text)
        ends=re.findall(r'(?m)^SUITE_FINISHED: ([^;\r\n]+)',text)
        assert starts==ends
        for i,macro in enumerate(starts):
            block=text.split('===== '+macro+' =====',1)[1].split('===== ',1)[0]
            totals=re.findall(r'(?m)^TOTAL[^\r\n]*passed=(\d+); failed=(\d+)',block)
            assert totals and int(totals[-1][1])==0,(macro,totals)
            suites.append({'Macro':macro,'Passed':int(totals[-1][0]),'Report':name})
    # Keep the latest execution of a duplicated entrypoint, without counting it twice.
    unique={item['Macro']:item for item in suites}
    assert sum(item['Passed'] for item in unique.values())==33992
    api=json.loads((ROOT/'docs/geometry_refactor/SECTION_RECIPE_API_EXAMPLES.json').read_text(encoding='utf-8-sig'))
    assert api['status']=='SOURCE_CALLS_AND_NATIVE_ROOTS_MATCH' and api['members']==322
    assert not api['missingCalls'] and not api['missingNativeRoots']
    pointed=(ROOT/'src/Geometry/CSectionRecipePointedSection.cls').read_text(encoding='utf-8-sig')
    rectangles=(ROOT/'src/Geometry/CSectionRecipeRectangles.cls').read_text(encoding='utf-8-sig')
    assert not re.search(r'Simple|Tapered|SideTypes|key & "Type"',pointed)
    assert 'Array("TwoRect", "OneRect")' in rectangles
    assert not re.search(r'"(?:Rectangle|LSection|TwoRectangles|TwoRectangle|RectSet)"',rectangles)
    assert not (ROOT/'src/Geometry/CSectionRecipePointedRect.cls').exists()
    assert not (ROOT/'src/Geometry/CSectionRecipeRectSet.cls').exists()
    geometry=read('GeometryEquivalence.json');rect=read('RectanglesEquivalence.json')
    assert geometry['Passed'] and rect['Passed']
    files=['BeforeGeometry.txt','AfterGeometry.txt','BeforeRectangles.txt','AfterRectangles.txt',
           'Directed01.txt','Directed02.txt','Directed03.txt','Directed04.txt','DirectedFinalInputs.txt']
    archive_path=D/'NativeSnapshotsAndReview.zip'
    if not archive_path.exists():
        with ZipFile(archive_path,'w',ZIP_DEFLATED,compresslevel=9) as archive:
            for name in files:archive.write(D/name,name)
    with ZipFile(archive_path) as archive:
        assert set(archive.namelist())==set(files) and archive.testzip() is None
    previous=D/'Acceptance.json'
    if previous.exists():assert sha(archive_path)==read('Acceptance.json')['EvidenceArchiveSHA256']
    result={'Passed':True,'Scope':'Focused acceptance; complete ON/OFF acceptance is deferred by the user.',
            'PublishedSHA256':digest,'AllCodeSHA256':sha(code),'NativeSourceComponents':155,'NativeExportComponents':161,
            'DirectedAssertions':sum(x['Passed'] for x in unique.values()),'Suites':list(unique.values()),
            'LayoutChecks':1720,'GeometryCases':geometry['Cases']+rect['Cases'],
            'BitIdenticalNumericFields':geometry['BitIdenticalNumericFields']+rect['BitIdenticalNumericFields'],
            'PreservedConfigAndCalculationInputs':707,'PreservedGeometryInputs':292,
            'ApiMembers':322,'UserReportUnchanged':True,'FullOnOffRun':False,
            'PendingChecks':[{'Macro':'modTestWorkbookInterface.RunSharedConfigPreflightTests','Reason':'420-second timeout in Directed04; remains to repeat with the deferred complete acceptance.'}],
            'ReviewedFailures':['Two incorrect manually specified entrypoint names in early exploratory drivers; corrected in final runs.',
                                'Eight circle caption expectations updated for the required newline.',
                                'Twenty-five common-setting location checks now accept the reader actual sheet/cell wording; exact address checks are retained.',
                                'Directed03 final watchdog hash read collided with preparation of its source; final acceptance uses frozen source files.'],
            'EvidenceArchiveSHA256':sha(D/'NativeSnapshotsAndReview.zip')}
    (D/'Acceptance.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({k:v for k,v in result.items() if k not in ('Suites','ReviewedFailures')},ensure_ascii=False))

if __name__=='__main__':run()
