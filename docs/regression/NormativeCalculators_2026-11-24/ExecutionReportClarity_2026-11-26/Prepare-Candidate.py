"""Готовит VBA-пакет отчета, сохраняя все пользовательские листы и результаты."""
import hashlib
import importlib.util
import json
from pathlib import Path
from zipfile import ZipFile

HERE = Path(__file__).resolve().parent
ROOT = next(p for p in HERE.parents if (p / 'AGENTS.md').is_file())
def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()
spec = importlib.util.spec_from_file_location('helpers', ROOT / 'tools/build_workbook/Package-NormativeCalculators.py')
h = importlib.util.module_from_spec(spec)
spec.loader.exec_module(h)
before, after = h.components(HERE / 'BaselineNative.txt'), h.components(HERE / 'NativeCandidate_Publication.txt')
assert before.keys() == after.keys() and len(after) == 209
expected = {'CExecutionReport', 'CBatchSectionCalculator', 'modWorkbookCalculation', 'CCapacityCalculator',
            'CCapacityResult', 'modTestWorkbookInterface', 'modTestBatchCalculation'}
changed = {name for name in before if h.normalized(before[name][1]) != h.normalized(after[name][1])}
assert changed == expected, changed
files = {p.stem: p for folder in ('src', 'tests') for p in (ROOT / folder).rglob('*') if p.suffix in ('.bas', '.cls')}
for name in expected:
    body = 'Option Explicit' + files[name].read_text(encoding='utf-8-sig').split('Option Explicit', 1)[1]
    if name == 'modTestWorkbookInterface':
        assert body.count('0.000000000000000001') == 2 and before[name][1].count('1E-18') >= 2
        body = body.replace('0.000000000000000001', '1E-18') # Нативная запись VBE, такая же в исходной книге; допуск не меняется.
    if h.normalized(body) != h.normalized(after[name][1]):
        import difflib
        diff = '\n'.join(difflib.unified_diff(h.normalized(body).splitlines(), h.normalized(after[name][1]).splitlines()))
        (HERE / (name + '_Difference.txt')).write_text(diff, encoding='utf-8')
        raise AssertionError(name)
with ZipFile(HERE / 'Baseline.xlsm') as baseline, ZipFile(HERE / 'CandidateNative.xlsm') as compiled, ZipFile(HERE / 'Candidate.xlsm', 'w') as output:
    for item in baseline.infolist():
        output.writestr(item, compiled.read(item.filename) if item.filename == 'xl/vbaProject.bin' else baseline.read(item.filename))
with ZipFile(HERE / 'Baseline.xlsm') as baseline, ZipFile(HERE / 'Candidate.xlsm') as candidate:
    assert baseline.namelist() == candidate.namelist()
    parts = [n for n in baseline.namelist() if baseline.read(n) != candidate.read(n)]
    assert parts == ['xl/vbaProject.bin']
dump = (HERE / 'NativeCandidate_Publication.txt').read_text(encoding='utf-8-sig')
dump = '\n'.join(line for line in dump.splitlines() if not line.startswith('Date:'))
dump = dump.replace('Workbook: CandidateNative.xlsm', 'Workbook: RC_Section_NDM.xlsm')
(HERE / 'VBA_All_Code.txt').write_bytes(b'\xef\xbb\xbf' + (dump.rstrip() + '\n').replace('\n', '\r\n').encode())
proof = {'Status': 'PASS', 'NativeComponents': 209, 'ChangedComponents': sorted(changed), 'SourceEquality': True,
         'BaselineSHA256': sha(HERE / 'Baseline.xlsm'), 'CandidateSHA256': sha(HERE / 'Candidate.xlsm'),
         'ExportSHA256': sha(HERE / 'VBA_All_Code.txt'), 'ChangedZIPParts': parts, 'AllUserDataAndFormattingPreserved': True}
(HERE / 'PackageProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps(proof, ensure_ascii=False))
