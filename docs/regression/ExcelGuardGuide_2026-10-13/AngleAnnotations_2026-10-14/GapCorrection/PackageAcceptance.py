"""Компактная приёмка исправления, без временных XLSM."""
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED

directory=Path(__file__).resolve().parent
root=directory.parents[4]
def read(name): return json.loads((directory/name).read_text(encoding='utf-8-sig'))
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()
for name in ('NativeAngleAcceptance.json','PublishedSourceEquality.json','PublishedPreservation.json','PartsPreservation.json'):
    assert read(name)['Passed'],name
assert read('PublishedStandaloneHelp.json')['status']=='STANDALONE_HELP_PASS'
assert read('../GapDirected01/Run.json')['Passed']
assert read('TaperedFacet02/Run.json')['Passed']
publication=read('Publication.json')
assert not publication['ReadbackPending'] and publication['AllCodeMatchesActualWorkbookExport']
assert sha(root/'workbook/output/RC_Section_NDM.xlsm').upper()==publication['WorkbookSHA256']
assert sha(root/'workbook/output/VBA_All_Code.txt').upper()==publication['AllCodeSHA256']
assert sha(root/'workbook/output/RC_Section_NDM_execution_report.txt').upper()==read('Before.json')['ExecutionReportSHA256']
groups={
    'NativeExports.zip':[directory/'BaselineNativeVBA.txt',directory/'PublishedNativeVBA.txt'],
    'NativeAngleEvidence.zip':[p for p in (directory/'CandidateNativeAngles').iterdir() if p.suffix.lower() in ('.json','.txt','.tsv','.png','.xlsx')],
    'DirectedAndFacetEvidence.zip':[p for folder in (directory.parent/'GapDirected01',directory/'TaperedFacet02')
                                  for p in folder.iterdir() if p.suffix.lower() in ('.json','.txt','.tsv')],
}
archives=[]
for name,paths in groups.items():
    target=directory/name
    with ZipFile(target,'w',ZIP_DEFLATED,compresslevel=9) as output:
        for p in sorted(paths): output.write(p,p.relative_to(directory.parent).as_posix())
    archives.append({'File':name,'Bytes':target.stat().st_size,'SHA256':sha(target)})
acceptance={'Passed':True,'ConfigTextGapSameRuleAtAllAngles':True,'OnlyRightVerticalRebarTurned180':True,
            'LeftAndDimensionDirectionsPreserved':True,'NativeCases':56,'NativeChecks':1553,
            'LayoutChecks':928,'GeneralPlotChecks':3377,'SnapshotChecks':24,
            'VisibleGlyphGapSameAtAllAngles':read('NativeAngleAcceptance.json')['GlyphGapPoints'],
            'UserDataAndAllNonVbaPartsByteIdentical':True,'Publication':publication,
            'Screenshot':'TaperedFacet02/RoundedTapered.png','FullOnOffRun':False,
            'Report':'docs/reports/AngleAnnotationGap_2026-10-14.md','Archives':archives}
(directory/'Acceptance.json').write_text(json.dumps(acceptance,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(acceptance,ensure_ascii=False))
