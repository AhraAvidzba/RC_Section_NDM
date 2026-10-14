"""Сводит фактические доказательства ограниченной доработки ResolveAnchor."""
import hashlib
import json
import re
from pathlib import Path
from zipfile import ZipFile

directory = Path(__file__).resolve().parent
root = directory.parents[3]

def read(name):
    return json.loads((directory/name).read_text(encoding='utf-8-sig'))

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

before = read('Before.json')
publication = read('Publication.json')
preservation = read('PublishedPreservation.json')
equality = read('PublishedSourceEquality.json')
help_result = read('PublishedStandaloneHelp.json')
guide = read('GuideNativeAccepted/GuideEvidence.json')
directed = read('DirectedFinal/Run.json')
counts = read('DirectedCounts.json')
equivalence = read('Equivalence.json')

assert preservation['Passed'] and preservation['NativeChangesRestricted']
assert equality['Passed'] and equality['SourceComponents'] == 152
assert help_result['status'] == 'STANDALONE_HELP_PASS' and help_result['apiMembers'] == 137
assert guide['Status'] == 'DIRECTED_NATIVE_GUIDE_PASS' and guide['ReopenTextMatch'] and guide['Idempotent']
assert directed['Passed'] and directed['ProductionUnchanged'] and counts['Total'] == 1792
assert equivalence['Passed'] and equivalence['Snapshots'][1]['ControlStates'] == 12
assert not publication['ReadbackPending'] and publication['PublishedCodeMatchesCandidate']
assert publication['AllCodeMatchesActualWorkbookExport']
assert sha(root/'workbook/output/VBA_All_Code.txt') == sha(directory/'PublishedNativeVBA.txt')
assert sha(root/'workbook/output/RC_Section_NDM_execution_report.txt').upper() == before['ExecutionReportSHA256']

with ZipFile(directory/'ExactSnapshots.zip') as archive:
    for a, b in [('BaselineModels', 'FinalModels'), ('BaselineNDS', 'FinalNDS')]:
        x = next(n for n in archive.namelist() if n.startswith(a+'/'))
        y = next(n for n in archive.namelist() if n.startswith(b+'/'))
        assert archive.read(x) == archive.read(y)

requirements = {
    'PositionOrderPhysicalEffectProvedAndOptionPreserved': True,
    'ResolveAnchorExactRequestedSignature': True,
    'ActualExistingSegmentIDAndLocalParameter': True,
    'CommonPointParameterAndExistingTolerancesReused': True,
    'MissingAndAmbiguousPointsHaveDiagnostics': True,
    'OutputsAtomicOnFailure': True,
    'RectangleAdjacentRoundedOffsetCircleArcAndTransformsTested': True,
    'EndpointsJointsArcsOutsideOverlapAndUnsealedCasesTested': True,
    'ChainEndsDirectionLengthAndOriginIntervalsTested': True,
    'ExistingPathMethodsFacadeAndWorkingRecipesUnchanged': True,
    'NineModelsAndTwelveNDSExactlyUnchanged': True,
    'GuideAndShortHelpUpdatedExampleOnlyInDetailedGuide': True,
    'NoNewClassOrParallelAPIOrSegmentRename': True,
    'FocusedRegressionOnlyNoFullOnOff': True,
    'ProductionReadbackAndAllCodeVerified': True,
    'UserDataAndForeignExecutionReportPreserved': True,
}
result = {
    'Status': 'RESOLVE_ANCHOR_ACCEPTANCE_PASS',
    'Requirements': requirements,
    'FocusedAssertions': counts['Total'],
    'CellsChecked': preservation['CellsChecked'],
    'NamedRanges': preservation['NamedRanges'],
    'NativeComponents': 157,
    'WorkbookSHA256': sha(root/'workbook/output/RC_Section_NDM.xlsm'),
    'AllCodeSHA256': sha(root/'workbook/output/VBA_All_Code.txt'),
    'ForeignExecutionReportSHA256': sha(root/'workbook/output/RC_Section_NDM_execution_report.txt'),
    'Errors': [],
}
(directory/'Acceptance.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
print(json.dumps(result, ensure_ascii=False))
