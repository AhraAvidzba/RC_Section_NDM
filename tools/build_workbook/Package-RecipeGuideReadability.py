"""Уточняет текст справочника с доказательством неизменности всего исполняемого VBA."""
import hashlib
import importlib.util
import json
import re
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
P = ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/RecipeSimplification_2026-10-16'
D = P/'GuideReadability'
spec = importlib.util.spec_from_file_location('presentation', Path(__file__).with_name('Verify-GuidePresentation.py'))
p = importlib.util.module_from_spec(spec)
spec.loader.exec_module(p)


def normalized(text, comments):
    return ''.join(part if part.startswith(('"', "'")) else re.sub(r'\s+', '', part.lower())
                   for part in re.findall(r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27]+', text)
                   if comments or not part.startswith("'"))


if __name__ == '__main__':
    a, b = [p.components(D/name) for name in ('BaselineNativeVBA.txt', 'CandidateNativeVBA.txt')]
    assert a.keys() == b.keys() and len(b) == 158
    assert all(normalized(a[name][2], False) == normalized(b[name][2], False) for name in a)
    changed = [name for name in a if normalized(a[name][2], True) != normalized(b[name][2], True)]
    assert changed == ['modSectionRecipeApiGuide'], changed
    assert all(not line.strip() or line.lstrip().startswith("'") for line in b['modSectionRecipeApiGuide'][2].splitlines())
    accepted = D/'Accepted.xlsm'
    assert not accepted.exists()
    with ZipFile(D/'Baseline.xlsm') as original, ZipFile(D/'Candidate.xlsm') as compiled, ZipFile(accepted, 'w') as result:
        for item in original.infolist():
            result.writestr(item, compiled.read(item.filename) if item.filename == 'xl/vbaProject.bin' else original.read(item.filename))
    with ZipFile(D/'Baseline.xlsm') as original, ZipFile(accepted) as result:
        assert original.namelist() == result.namelist()
        assert [name for name in original.namelist() if original.read(name) != result.read(name)] == ['xl/vbaProject.bin']
    sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
    before = json.loads((D/'Before.json').read_text(encoding='utf-8-sig'))
    assert sha(D/'Baseline.xlsm') == before['WorkbookSHA256'].lower()
    acceptance = json.loads((P/'Acceptance.json').read_text(encoding='utf-8-sig'))
    assert acceptance['Status'] == 'RECIPE_SIMPLIFICATION_PASS' and acceptance['PublishedOutputVerified']
    acceptance['InheritedFromUnchangedExecutableCode'] = True
    (D/'Acceptance.json').write_text(json.dumps(acceptance, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    receipt = {'Status': 'RECIPE_PACKAGE_PASS', 'AcceptedSHA256': sha(accepted),
               'BeforeWorkbookSHA256': sha(D/'Baseline.xlsm'), 'NativeComponents': 158,
               'AllExecutableVBAUnchanged': True, 'AllNonVBAPartsByteIdentical': True,
               'ChangedCommentComponents': changed, 'OnlyZIPPartChanged': 'xl/vbaProject.bin'}
    (D/'Package.json').write_text(json.dumps(receipt, indent=2)+'\n', encoding='utf-8')
    print(json.dumps(receipt))
