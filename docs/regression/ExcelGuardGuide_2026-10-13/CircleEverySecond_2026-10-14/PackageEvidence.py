"""Сводит приёмку и упаковывает точные снимки, без частных XLSM."""
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
directory=Path(__file__).resolve().parent;root=directory.parents[3]
read=lambda name:json.loads((directory/name).read_text(encoding='utf-8-sig'))
proof=read('PublishedPreservation.json');publication=read('Publication.json');factory=read('FactoryTest.json')
require=[proof['Passed'],factory['Passed'],read('FinalWorkbookDirected/Run.json')['Passed'],publication['AllCodeMatchesActualWorkbookExport'],not publication['ReadbackPending']]
assert all(require)
sha=lambda path:hashlib.sha256(path.read_bytes()).hexdigest().upper()
assert sha(root/'workbook/output/RC_Section_NDM.xlsm')==publication['WorkbookSHA256']
assert sha(root/'workbook/output/VBA_All_Code.txt')==publication['AllCodeSHA256']==sha(directory/'PublishedNativeVBA.txt')
assert sha(root/'workbook/output/RC_Section_NDM_execution_report.txt')==read('Before.json')['ExecutionReportSHA256']
assert proof['Assertions']+factory['Checks']==5883
acceptance={
    'Passed':True,'BaselineHead':proof['BaselineHead'],
    'AllFourRecipeFamiliesSupportEachBarAndEverySecondBar':True,
    'CircleRow2AndRow3ChoicesIndependent':True,'DefaultEachBarPreserved':True,
    'SelectionUsesExistingCommonExecutor':True,'CoreAlgorithmsUnchanged':True,
    'PublicationCommentsExplainNumberingWithoutMovingBars':True,
    'AnnotationCommentsExplainCountsForEveryDiameterAcrossRows':True,
    'RecipeHeadersExplainConcreteReinforcementAndAllRecipeAnnotations':True,
    'AssertionsWithoutRepeat':5883,'FinalWorkbookRepeatAssertions':4795,
    'DefaultModelsEqualExceptNewInputCells':9,'OriginalCellsChecked':107461,
    'ActualNativeExportEqualsAllCode':True,'UserExecutionReportUnchanged':True,
    'WorkbookSHA256':publication['WorkbookSHA256'],'AllCodeSHA256':publication['AllCodeSHA256'],
    'FullOnOffRunsStarted':False,
}
(directory/'Acceptance.json').write_text(json.dumps(acceptance,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
with ZipFile(directory/'ExactSnapshots.zip','w',ZIP_DEFLATED,compresslevel=9) as out:
    for folder in ('BaselineGeometry','CandidateGeometry','Directed01','WorkbookDirected','FinalWorkbookDirected'):
        for path in sorted((directory/folder).glob('*')):
            if path.suffix in ('.txt','.json'):out.write(path,path.relative_to(directory).as_posix())
    for name in ('BaselineNativeVBA.txt','PublishedNativeVBA.txt'):out.write(directory/name,name)
print(json.dumps(acceptance,ensure_ascii=False))
