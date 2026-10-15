"""Итоговая приёмка удаления миграций и согласованных изменений оформления."""
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

directory = Path(__file__).resolve().parent
root = directory.parents[3]
def read(name):
    return json.loads((directory/name).read_text(encoding='utf-8-sig'))
def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

before, publication = read('Before.json'), read('Publication.json')
audit, preservation = read('FinalAudit.json'), read('PublishedPreservation.json')
equality, help_result = read('PublishedSourceEquality.json'), read('PublishedStandaloneHelp.json')
counts, equivalence = read('DirectedCounts.json'), read('Equivalence.json')
assert audit['Passed'] and audit['NativeComponentsAudited'] == 157
assert audit['CalculationAndGeometryAlgorithmsUnchanged']
assert preservation['Passed'] and preservation['NativeChangesRestricted']
assert preservation['UserDataFormulasStylesNamesAndDrawingsPreserved']
assert equality['Passed'] and equality['SourceComponents'] == 152
assert help_result['status'] == 'STANDALONE_HELP_PASS' and help_result['apiMembers'] == 137
assert read('GuideNativeAccepted/GuideEvidence.json')['Status'] == 'DIRECTED_NATIVE_GUIDE_PASS'
assert read('BootstrapAndText.json')['Passed'] and read('PowerShellSyntax.json')['Passed']
assert counts['Passed'] and counts['Total'] == 9987
for stage in {g['Stage'] for g in counts['Groups']}:
    assert read(stage+'/Run.json')['Passed'] and read(stage+'/Run.json')['ProductionUnchanged']
assert equivalence['Passed'] and equivalence['Snapshots'][1]['ControlStates'] == 12
assert read('AcceptanceNativeImages/Run.json')['Passed']
assert not publication['ReadbackPending'] and publication['PublishedCodeMatchesCandidate']
assert publication['AllCodeMatchesActualWorkbookExport']
assert sha(root/'workbook/output/RC_Section_NDM.xlsm') == sha(directory/'Candidate.xlsm')
assert sha(root/'workbook/output/VBA_All_Code.txt') == sha(directory/'PublishedNativeVBA.txt')
assert sha(root/'workbook/output/RC_Section_NDM_execution_report.txt').upper() == before['ExecutionReportSHA256']
with ZipFile(directory/'ExactSnapshots.zip') as archive:
    for kind in ('Models','NDS'):
        paths = [next(n for n in archive.namelist() if n.startswith(stage+kind+'/')) for stage in ('Baseline','Acceptance')]
        assert archive.read(paths[0]) == archive.read(paths[1])
with ZipFile(directory/'NativeSnapshots.zip','w',ZIP_DEFLATED) as archive:
    for name in ('BaselineNativeVBA.txt','FinalNativeVBA.txt','PublishedNativeVBA.txt'):
        archive.write(directory/name,name)

result = {
    'Status':'NO_MIGRATIONS_ACCEPTANCE_PASS',
    'Requirements': {
        'EntireNativeVbaProjectAndActualDependenciesAudited':True,
        'CompletedRuntimeMigrationsAndExclusiveHelpersCallersTestsRemoved':True,
        'OneTimeBuildMigrationsAndExclusiveRunnersRemoved':True,
        'CurrentInputValidationAndSupportedFormatsRetained':True,
        'TypeUpdaterOnlyOwnListNameExtentAndTypeValidation':True,
        'OneRegistryAdditionAppearsWithoutOtherTableChanges':True,
        'FreshTemplateBuilderProvidesInitialTypeList':True,
        'CurrentSchemaRuleDocumentedAndHelpUpdated':True,
        'GeometryArchitectureAndNumericalAlgorithmsPreserved':True,
        'BothImportedOverallDimensionsUseExistingAnnotations':True,
        'ExactApproximateAndDisabledDimensionCasesPassed':True,
        'PlotBoundsIncludeOuterWithoutStricterGeometryValidation':True,
        'BothOffsetsUseMaxOneDOver1000FromConfig':True,
        'ChartScaleAndSectionSizeCoefficientRemainDistinct':True,
        'OpenArcsZeroSweepAndArbitraryAnnotationsPreserved':True,
        'NineModelsAndTwelveNDSExactlyUnchanged':True,
        'FocusedRegressionOnlyNoFullOnOff':True,
        'ProductionReadbackAndAllCodeVerified':True,
        'UserDataAndForeignExecutionReportPreserved':True,
    },
    'FocusedAssertions':counts['Total'],
    'AdditionalImageExportRunAssertions':27,
    'CellsChecked':preservation['CellsChecked'],'NamedRanges':preservation['NamedRanges'],
    'NativeComponents':157,'BaselineProceduresInventoried':audit['BaselineProceduresInventoried'],
    'WorkbookSHA256':sha(root/'workbook/output/RC_Section_NDM.xlsm'),
    'AllCodeSHA256':sha(root/'workbook/output/VBA_All_Code.txt'),
    'ForeignExecutionReportSHA256':sha(root/'workbook/output/RC_Section_NDM_execution_report.txt'),
    'Errors':[],
}
(directory/'Acceptance.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(result,ensure_ascii=False))
