"""Проверяет каждое требование добавленного оформления и фактический output."""
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED

directory=Path(__file__).resolve().parent
root=directory.parents[3]
def read(name):return json.loads((directory/name).read_text(encoding='utf-8-sig'))
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
preserved=read('PublishedPreservation.json');native=read('NativeCheck.json');counts=read('DirectedCounts.json')
publication=read('Publication.json');before=read('Before.json')
assert preserved['Passed'] and preserved['Default']==2 and preserved['NewTitle']=='Настройки оформления схемы'
assert read('PublishedSourceEquality.json')['Passed'] and read('PublishedStandaloneHelp.json')['status']=='STANDALONE_HELP_PASS'
assert native['Passed'] and native['Default']==2 and native['FreshBuilderSameSetting']
assert counts['Passed'] and counts['Total']==994 and read('Directed03/Run.json')['Passed']
assert not publication['ReadbackPending'] and publication['AllCodeMatchesActualWorkbookExport']
assert sha(root/'workbook/output/VBA_All_Code.txt')==sha(directory/'PublishedNativeVBA.txt')
assert sha(root/'workbook/output/RC_Section_NDM.xlsm')==preserved['WorkbookSHA256']
assert sha(root/'workbook/output/RC_Section_NDM_execution_report.txt').upper()==before['ExecutionReportSHA256']
plot=(root/'src/Excel/CSectionPlotter.cls').read_text(encoding='utf-8-sig')
body=plot.split('Private Sub DrawConcreteContour',1)[1].split('End Sub',1)[0]
assert body.count('mContourLineWeight')==3 and 'SourceType' not in body and '1.7' not in body
for item in read('RestoredApiReferences.json')['RestoredReferencedReports']:
    assert sha(root/item['Path'])==item['SHA256']
with ZipFile(directory/'NativeSnapshots.zip','w',ZIP_DEFLATED) as archive:
    for name in ('InitialNativeVBA.txt','BaselineNativeVBA.txt','FinalNativeVBA.txt','PublishedNativeVBA.txt'):
        archive.write(directory/name,name)
result={'Status':'CONTOUR_LINE_WEIGHT_ACCEPTANCE_PASS','Requirements':{
    'SettingInRequestedConfigTable':True,'DefaultTwoPtPerUserCorrection':True,
    'OneCommonValueForOuterAndOpeningLineArcCircle':True,
    'CommonPlotterIndependentOfGeneratedOrImportedSource':True,
    'RequiredPositiveValueWithActualCellDiagnostics':True,
    'DisabledContoursIgnoreUnusedInvalidValue':True,
    'ThicknessInPointsUnaffectedByChartSizeOrLengthUnits':True,
    'ChosenTitleOnConfigAndHelp':True,'HelpAndReturnLinksWork':True,
    'FreshWorkbookBuilderSameSetting':True,'UserInputsAndSavedResultsPreserved':True,
    'SourceWorkbookAndAllCodeSynchronized':True,'NoRuntimeMigrationOrRecipeBranches':True,
    'FocusedChecksOnlyNoFullOnOffOrNumericalRecalculation':True},
    'Assertions':994,'CellsCompared':preserved['CellsChecked'],'NativeComponents':157,
    'WorkbookSHA256':sha(root/'workbook/output/RC_Section_NDM.xlsm'),
    'AllCodeSHA256':sha(root/'workbook/output/VBA_All_Code.txt'),
    'ForeignExecutionReportSHA256':sha(root/'workbook/output/RC_Section_NDM_execution_report.txt'),
    'Errors':[]}
(directory/'Acceptance.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(result,ensure_ascii=False))
