"""Готовит принятую структуру книги с обновлёнными VBA и одним абзацем справки."""
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
from xml.sax.saxutils import escape
from xml.etree import ElementTree as ET
from zipfile import ZipFile

HERE = Path(__file__).resolve().parent
ROOT = next(p for p in HERE.parents if (p / 'AGENTS.md').is_file())


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


spec = importlib.util.spec_from_file_location('package_helpers', ROOT / 'tools/build_workbook/Package-NormativeCalculators.py')
helpers = importlib.util.module_from_spec(spec)
spec.loader.exec_module(helpers)
before = helpers.components(HERE / 'BaselineNative.txt')
after = helpers.components(HERE / 'NativeCandidate.txt')
assert len(before) == len(after) == 209 and before.keys() == after.keys()
changed = {name for name in before if helpers.normalized(before[name][1]) != helpers.normalized(after[name][1])}
expected = {'CTorsionSummaryWriter', 'modCalculationSnapshot', 'modTestTorsionPost', 'modTestTorsionConcrete'}
assert changed == expected, changed
files = {p.stem: p for folder in ('src', 'tests') for p in (ROOT / folder).rglob('*') if p.suffix in ('.bas', '.cls')}
for name in expected:
    body = files[name].read_text(encoding='utf-8-sig').split('Option Explicit', 1)[1]
    assert helpers.normalized('Option Explicit' + body) == helpers.normalized(after[name][1]), name
old_catalog = json.loads(subprocess.check_output(['git', 'show', 'e80b5f8476d1aa9967a0c2f41fb488abf0165f09:tools/build_workbook/HelpSettingsDescriptions.json'], cwd=ROOT).decode('utf-8-sig'))
new_catalog = json.loads((ROOT / 'tools/build_workbook/HelpSettingsDescriptions.json').read_text(encoding='utf-8-sig'))
old = next(x for x in old_catalog if x['Title'] == 'Results.SnapshotStorage')['Lines'][2]
new = next(x for x in new_catalog if x['Title'] == 'Results.SnapshotStorage')['Lines'][2]
assert len(new) <= len(old)
old_bytes, new_bytes = escape(old).encode(), escape(new).encode()
target = HERE / 'Candidate.xlsm'
with ZipFile(HERE / 'Baseline.xlsm') as baseline, ZipFile(HERE / 'CandidateNative.xlsm') as compiled:
    parts = [name for name in baseline.namelist() if name.endswith('.xml') and baseline.read(name).count(old_bytes)]
    assert len(parts) == 1, parts
    help_part = parts[0]
    assert baseline.read(help_part).count(old_bytes) == 1
    help_bytes = baseline.read(help_part).replace(old_bytes, new_bytes)
    ET.fromstring(help_bytes)
    with ZipFile(target, 'w') as output:
        for item in baseline.infolist():
            data = compiled.read(item.filename) if item.filename == 'xl/vbaProject.bin' else help_bytes if item.filename == help_part else baseline.read(item.filename)
            output.writestr(item, data)
with ZipFile(HERE / 'Baseline.xlsm') as baseline, ZipFile(target) as candidate:
    assert baseline.namelist() == candidate.namelist()
    parts_changed = [n for n in baseline.namelist() if baseline.read(n) != candidate.read(n)]
    assert set(parts_changed) == {'xl/vbaProject.bin', help_part}
text = (HERE / 'NativeCandidate.txt').read_text(encoding='utf-8-sig')
text = '\n'.join(line for line in text.splitlines() if not line.startswith('Date:'))
text = text.replace('Workbook: CandidateNative.xlsm', 'Workbook: RC_Section_NDM.xlsm')
(HERE / 'VBA_All_Code.txt').write_bytes(b'\xef\xbb\xbf' + (text.rstrip() + '\n').replace('\n', '\r\n').encode())
proof = {'Status': 'PASS', 'BaselineSHA256': sha(HERE / 'Baseline.xlsm'), 'CandidateSHA256': sha(target),
         'ExportSHA256': sha(HERE / 'VBA_All_Code.txt'), 'CompiledNativeSHA256': sha(HERE / 'CandidateNative.xlsm'),
         'NativeComponents': len(after), 'ChangedComponents': sorted(changed), 'SourceEquality': True,
         'ChangedZIPParts': parts_changed, 'OnlyOneHelpParagraphChanged': True, 'HelpTextNotLonger': True,
         'AllOtherWorkbookPartsPreserved': True}
(HERE / 'PackageProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps(proof, ensure_ascii=False))
