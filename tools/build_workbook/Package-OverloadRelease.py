"""Переносит только проверенный VBA в последнюю пользовательскую книгу."""
import hashlib
import importlib.util
import json
import re
from pathlib import Path
from zipfile import ZipFile

root = Path(__file__).resolve().parents[2]
directory = root / 'docs/regression/Performance/Overload_2026-10-15/Release'
spec = importlib.util.spec_from_file_location('presentation', root / 'tools/build_workbook/Verify-GuidePresentation.py')
p = importlib.util.module_from_spec(spec)
spec.loader.exec_module(p)


def normalize(text):
    return ''.join(part if part.startswith(('"', "'")) else part.lower() for part in re.findall(
        r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27\n]+|\n', text.replace('\r\n', '\n')))


baseline = p.components(directory / 'BaselineNativeVBA.txt')
candidate = p.components(directory / 'CandidateNativeVBA.txt')
all_code = p.components(root / 'workbook/output/VBA_All_Code.txt')
assert baseline.keys() == candidate.keys() == all_code.keys() and len(candidate) == 157
assert all(normalize(baseline[name][2]) == normalize(all_code[name][2]) for name in baseline), 'User VBA differs from current ALL CODE; reconcile it before publishing.'
changed_components = sorted(name for name in baseline if normalize(baseline[name][2]) != normalize(candidate[name][2]))
assert set(changed_components) == {'CCapacitySolver', 'CCapacityLimitSearchProblem', 'CLimitSearchCoordinator', 'CUltimateStrainSearch', 'modTestCapacitySolver'}
assert 'RunOverloadTimingCase' not in candidate['modTestBatchCalculation'][2]
assert 'CaptureFocusedCalculation' not in candidate['modWorkbookCalculation'][2]
accepted = directory / 'Accepted.xlsm'
assert not accepted.exists()
with ZipFile(directory / 'Baseline.xlsm') as a, ZipFile(directory / 'Candidate/RC_Section_NDM.xlsm') as b, ZipFile(accepted, 'w') as output:
    for entry in a.infolist():
        output.writestr(entry, b.read(entry.filename) if entry.filename == 'xl/vbaProject.bin' else a.read(entry.filename))
with ZipFile(directory / 'Baseline.xlsm') as a, ZipFile(accepted) as b:
    assert a.namelist() == b.namelist()
    changed = [name for name in a.namelist() if a.read(name) != b.read(name)]
    assert changed == ['xl/vbaProject.bin']
receipt = {
    'Status': 'OVERLOAD_RELEASE_PACKAGE_PASS', 'NativeComponents': 157,
    'ChangedNativeComponents': changed_components, 'ChangedZIPParts': changed,
    'AllCellsFormulasSettingsNamesValidationFormattingDrawingsPreserved': True,
    'NoMeasurementHelpersInProduction': True,
    'BaselineSHA256': hashlib.sha256((directory / 'Baseline.xlsm').read_bytes()).hexdigest(),
    'WorkbookSHA256': hashlib.sha256(accepted.read_bytes()).hexdigest(),
    'Errors': [],
}
(directory / 'Package.json').write_text(json.dumps(receipt, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps(receipt, ensure_ascii=False))
