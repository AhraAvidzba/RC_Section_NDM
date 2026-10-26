"""ДЛЯ ТЕСТОВ: заменяет только VBA в проверенном пакете с сохранёнными данными."""
from pathlib import Path
import argparse, hashlib, json, zipfile

here = Path(__file__).resolve().parent
parser = argparse.ArgumentParser()
parser.add_argument('--compiled', required=True)
parser.add_argument('--compiled-sha256', required=True)
parser.add_argument('--label', required=True)
args = parser.parse_args()
source = here / 'StyleFixedRelease03/RC_Section_NDM.xlsm'
compiled = here / args.compiled
directory = here / args.label
assert directory.parent == here and compiled.resolve().is_relative_to(here.resolve())
directory.mkdir(exist_ok=False)
output = directory / 'RC_Section_NDM.xlsm'
digest = lambda p: hashlib.sha256(p.read_bytes()).hexdigest().upper()
assert digest(source) == '27625E01849E32161205757182112E2A8386B1B19CD5000E2FF482B33AAD5E60'
assert digest(compiled) == args.compiled_sha256.upper()
with zipfile.ZipFile(source) as base, zipfile.ZipFile(compiled) as code:
    assert base.testzip() is None and code.testzip() is None
    with zipfile.ZipFile(output, 'w') as target:
        for part in base.infolist():
            target.writestr(part, code.read(part.filename) if part.filename == 'xl/vbaProject.bin' else base.read(part.filename))
with zipfile.ZipFile(source) as base, zipfile.ZipFile(output) as target:
    assert target.testzip() is None and base.namelist() == target.namelist()
    changes = [name for name in base.namelist() if base.read(name) != target.read(name)]
    assert changes == ['xl/vbaProject.bin']
    proof = dict(Status='PREPARED_REQUIRES_FINAL_GATES', BaseSHA256=digest(source),
                 CompiledSHA256=digest(compiled), WorkbookSHA256=digest(output),
                 Parts=len(base.namelist()), ChangedParts=changes,
                 UnchangedParts=len(base.namelist())-1, FullRegressionPass=False)
(directory/'PackageProof.json').write_text(json.dumps(proof, indent=2), encoding='utf-8')
print(json.dumps(proof))
