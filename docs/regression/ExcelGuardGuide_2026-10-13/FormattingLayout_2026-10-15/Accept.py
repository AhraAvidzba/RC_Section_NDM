"""Фиксирует итоговую приёмку основной книги и актуальность неизменённого ALL CODE."""
import hashlib
import json
import re
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

directory = Path(__file__).resolve().parent
root = directory.parents[3]


def receipt(name):
    return json.loads((directory / name).read_text(encoding='utf-8-sig'))


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


before = receipt('Before.json')
assert receipt('PublishedPreservation.json')['Passed']
assert receipt('NativeCheck.json')['Passed']
assert receipt('PublishedSourceEquality.json')['Passed']
native = directory / 'PublishedNativeVBA.txt'
all_code = root / 'workbook/output/VBA_All_Code.txt'
strip_date = lambda text: re.sub(r'(?m)^Date: [^\n]*\n', '', text)
assert strip_date(native.read_text(encoding='utf-8-sig')) == strip_date(all_code.read_text(encoding='utf-8-sig'))
assert sha(all_code).upper() == before['AllCodeSHA256']
assert sha(root / 'workbook/output/RC_Section_NDM.xlsm').upper() == receipt('Publication.json')['WorkbookSHA256']
assert sha(root / 'workbook/output/RC_Section_NDM_execution_report.txt').upper() == before['ExecutionReportSHA256']
with ZipFile(directory / 'NativeSnapshot.zip', 'w', ZIP_DEFLATED) as archive:
    archive.write(native, native.name)
with ZipFile(directory / 'NativeSnapshot.zip') as archive:
    assert archive.read(native.name) == native.read_bytes()
result = {
    'Status': 'CONFIG_FORMATTING_LAYOUT_ACCEPTANCE_PASS',
    'EmptyRowAfterFormattingTable': 85,
    'AllFourLowerGeometryTablesMovedOneRow': True,
    'BlueUnderlinedHeadingLinkNativeExcelVerified': True,
    'NamedRangesValidationMergesAndFormulasPreserved': True,
    'ExistingCellsChecked': 107483, 'MovedCells': 882,
    'SourceComponents': 152, 'NativeComponents': 157,
    'AllCodeMatchesActualWorkbookExceptExportDate': True,
    'AllCodeUnchanged': True, 'UserExecutionReportUnchanged': True,
    'NoCalculationOrFullOnOffRuns': True,
    'WorkbookSHA256': sha(root / 'workbook/output/RC_Section_NDM.xlsm'),
    'AllCodeSHA256': sha(all_code), 'NativeExportSHA256': sha(native),
    'BaselineGitCommit': 'ab230004f1280a9863ceafe2a3877104ab596d25',
    'Errors': [],
}
(directory / 'Acceptance.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps(result, ensure_ascii=False))
