"""Финальная приёмка уточнённой цели по сохранённым доказательствам."""
import hashlib
import json
from pathlib import Path

directory = Path(__file__).resolve().parent
root = directory.parents[3]
read = lambda name: json.loads((directory/name).read_text(encoding='utf-8-sig'))
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
style, geometry, examples = read('RecipeStyleAcceptance.json'), read('NativeGeometryAcceptance.json'), read('CalculatedExamplesAcceptance.json')
preservation, equality, publication = read('PublishedPreservation.json'), read('PublishedSourceEquality.json'), read('Publication.json')
image = read('RecipeCircleImage.json')
checks = {
    '1_PlainRussianRecipeHeadersAnd30MethodComments': style['Passed'] and sum(c['MethodsReviewed'] for c in style['Classes']) == 30,
    '2_InlineShortCommentsAndNoUnnecessaryShortContinuations': all(c['ShortCommentsInline'] and c['UnnecessaryShortContinuations'] == 0 for c in style['Classes']),
    '3_EstablishedHeaderStyle': style['ReferenceStyle'] == ['CSectionModel', 'CMaterialDiagram'],
    '4_FewerObjectsWithPreservedFreeformAxesAndGeometry': geometry['Passed'] and geometry['Cases'] == 162 and geometry['Checks'] == 8649,
    'DirectedPlotChecks': read('CandidateDirected/Run.json')['Passed'],
    'ActualGConcreteAndTransformedAxesRelativeToGeometryAndLoad': examples['Passed'] and len(examples['Cases']) == 7,
    'GCalculatedCoordinatesStressesAndStatusUnchanged': examples['BaselineGCalculationUnchanged'],
    'NativeScreenshotsOfCalculatedGAndRoundedHollow': read('CalculatedExamples.json')['Passed'],
    'CircleCodeImageMatchesPublishedSource': image['TextFromActualSource'] and image['NoCodeOmitted'] and image['SourceSHA256'] == sha(root/'src/Geometry/CSectionRecipeCircle.cls'),
    'UserDataResultsFormulasStylesHelpPreserved': preservation['Passed'] and preservation['InputResultsFormulasStylesAndHelpUnchanged'] and preservation['CellsChecked'] == 107461 and preservation['NamedRanges'] == 27,
    'All152SourceComponentsMatchActualWorkbook': equality['Passed'],
    'NoTemporaryCodeOrEmbeddedMarkers': preservation['TemporaryHelpersAbsent'] and preservation['NoEmbeddedMarkers'],
    'FinalWorkbookAndAllCodePublishedAndReadBack': not publication['ReadbackPending'] and publication['AllCodeMatchesActualWorkbookExport'],
    'PublishedHashesStillMatch': sha(root/'workbook/output/RC_Section_NDM.xlsm').upper() == publication['WorkbookSHA256'] and sha(root/'workbook/output/VBA_All_Code.txt').upper() == publication['AllCodeSHA256'],
    'UserExecutionReportUnchanged': sha(root/'workbook/output/RC_Section_NDM_execution_report.txt').upper() == read('Before.json')['ExecutionReportSHA256'],
    'AllCodeIsActualExportBytes': (root/'workbook/output/VBA_All_Code.txt').read_bytes() == (directory/'PublishedNativeVBA.txt').read_bytes(),
}
result = {'Passed': all(checks.values()), 'Requirements': checks, 'FullOnOffOrLongUIPerformed': False, 'Scope': 'Recipe comments, plot object minimization, actual calculated examples and coordinate acceptance', 'Errors': [name for name, passed in checks.items() if not passed]}
(directory/'Acceptance.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
print(json.dumps(result, ensure_ascii=False))
if not result['Passed']:
    raise SystemExit(1)
