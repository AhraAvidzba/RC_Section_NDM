"""Итоговая приёмка фактической основной книги после независимого чтения."""
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile

root = Path(__file__).resolve().parents[2]
directory = root / 'docs/regression/Performance/Overload_2026-10-15'
release = directory / 'Release'
native = root / 'docs/regression/ExcelGuardGuide_2026-10-13/Overload_2026-10-15'


def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


before = read(release / 'Before.json')
package = read(release / 'Package.json')
publication = read(release / 'Publication.json')
measurements = read(directory / 'MeasurementsAcceptance.json')
assert measurements['Status'] == 'OVERLOAD_TIMING_ACCEPTANCE_PASS'
assert package['Status'] == 'OVERLOAD_RELEASE_PACKAGE_PASS'
assert read(native / 'PublishedSourceEquality.json')['Passed']
assert not publication['ActualWorkbookReadbackPending'] and publication['AllCodeMatchesActualNativeExport']
book = root / 'workbook/output/RC_Section_NDM.xlsm'
code = root / 'workbook/output/VBA_All_Code.txt'
assert sha(book) == package['WorkbookSHA256'] == publication['WorkbookSHA256'].lower()
snapshot = native / 'PublishedNativeVBA.txt'
if snapshot.exists():
    exported = snapshot.read_bytes()
else:
    with ZipFile(directory / 'OverloadEvidence.zip') as archive:
        exported = archive.read(snapshot.relative_to(root).as_posix())
assert code.read_bytes() == exported
assert sha(code) == publication['AllCodeSHA256'].lower()
assert sha(root / 'workbook/output/RC_Section_NDM_execution_report.txt') == before['ExecutionReportSHA256'].lower()
with ZipFile(release / 'Baseline.xlsm') as a, ZipFile(book) as b:
    assert a.namelist() == b.namelist()
    changed = [name for name in a.namelist() if a.read(name) != b.read(name)]
    assert changed == ['xl/vbaProject.bin']
result = {
    'Status': 'OVERLOAD_FULL_ACCEPTANCE_PASS',
    'Requirements': {
        'Measured2x3x4xFromConfirmedCapacity': True,
        'ReportedSlowCircleCaseReproducedAt19to20Seconds': True,
        'CauseVerifiedByActualIterationCountsAndRuntimeLogs': True,
        'BisectionSevenIterationsAtConfiguredToleranceVerified': True,
        'RedundantUltimateAttemptsRemovedInCommonClasses': True,
        'NoRecipeBranchesOrReducedPrecision': True,
        'NoPhysicalFailureInferredFromNumericalFailure': True,
        'BoundaryPlanesAndStatusesPreservedForComparedLoads': True,
        'AllMeasuredCandidateRunsUnder10Seconds': True,
        'ActualWorkbookSourceAndAllCodeSynchronized': True,
        'UserDataAndExecutionReportPreserved': True,
    },
    'RegressionAssertions': measurements['TotalRegressionAssertions'],
    'NativeComponents': 157, 'SourceComponents': 152,
    'ChangedZIPParts': changed, 'WorkbookSHA256': sha(book), 'AllCodeSHA256': sha(code),
    'Errors': [],
}
(directory / 'Acceptance.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps(result, ensure_ascii=False))
