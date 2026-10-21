"""Archive measured evidence; assert completeness without declaring all variants PASS."""
from pathlib import Path
import csv
import hashlib
import json
import zipfile

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()

expected_sources = {
    'workbook/output/RC_Section_NDM.xlsm': '7B4C769D3813CE0057FFA93026281CD8ED7D3DC23F60159A2655882AB37D2772',
    'workbook/output/VBA_All_Code.txt': 'F88F3868FF9C3FF0D14DB981D6F90CC1467D4CB0E426A462497EF86FC158950E',
    'workbook/output/RC_Section_NDM_execution_report.txt': 'A973C4D3FBA2E2042E1BE25E62B27462C6C44A4749CB0D749B646C48E2252377',
}
for name, expected in expected_sources.items():
    assert sha(ROOT / name) == expected, f'Changed outside experiment: {name}'

summary = json.loads((HERE / 'Summary.json').read_text(encoding='utf-8'))
plans = {
    'Light_Full': (3, 'Full NoPlot Outline Grouped Plain NoPlotPlain NoJson Minimum Memory NoReport'),
    'Dense_Full': (2, 'Full NoPlot Outline Grouped Plain NoPlotPlain NoJson Minimum Memory NoReport FullEnd'),
    'Light_JSONExact': (3, 'Full JsonExact JsonExactNoPlot NoPlot FullEnd'),
    'Dense_JSONExact': (2, 'Full JsonExact JsonExactNoPlot NoPlot FullEnd'),
    'Batch_Core': (2, 'Full NoPlot Minimum Memory'),
    'Accuracy_Checks': (3, 'Full Mesh20 Mesh40 NoCapacity NoCracks NoStability DirectOnly FullEnd'),
}
verified = {}
for name, (repetitions, names) in plans.items():
    names = names.split()
    run = summary[name]
    assert run['completed'] and not run['failuresReported'], name
    assert set(run['variants']) == set(names), name
    rows = list(csv.DictReader((HERE / (name + '.csv')).read_text(encoding='utf-8-sig').splitlines()))
    assert len(rows) == len(names) * (repetitions + 1), name
    for variant in names:
        samples = [r for r in rows if r['variant'] == variant]
        assert sorted(int(r['repetition']) for r in samples) == list(range(repetitions + 1))
        if name != 'Accuracy_Checks' and variant != 'NoReport':
            assert all(int(r['mismatch']) == 0 for r in samples), (name, variant)
    verified[name] = {'measuredRepetitions': repetitions, 'variants': names, 'rowsIncludingWarmups': len(rows)}

files = []
for path in HERE.iterdir():
    if not path.is_file() or path.name in ['EvidenceIndex.json', 'Modules.json']:
        continue
    if path.suffix in ['.csv', '.tsv', '.json'] or (path.suffix == '.txt' and not path.name.endswith('.bas.txt')):
        files.append(path)
files.extend(HERE.glob('Candidate*/Preparation.json'))
files = sorted(files, key=lambda p: p.relative_to(HERE).as_posix())
manifest = []
archive = HERE / 'Evidence.zip'
with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as bundle:
    for path in files:
        name = path.relative_to(HERE).as_posix()
        bundle.write(path, name)
        item = {'path': name, 'bytes': path.stat().st_size, 'sha256': sha(path)}
        if path.name.endswith('_messages.txt'):
            item['encoding'] = 'Windows-1251 (original VBA Print output)'
        manifest.append(item)
with zipfile.ZipFile(archive) as bundle:
    assert bundle.testzip() is None
    for item in manifest:
        assert hashlib.sha256(bundle.read(item['path'])).hexdigest().upper() == item['sha256']

index = {
    'scope': 'Performance tradeoffs in isolated copies; not production acceptance or a full ON/OFF run',
    'baseCommit': '4ccfe9c7336d4c49e665dee42a4e5d24c200c1ff',
    'branch': 'codex/linked-rebar-row-offsets',
    'productionUnchanged': expected_sources,
    'verifiedCompletedPlans': verified,
    'exceptions': [
        'NoReport: one diagnostic LastResidualNorm changes; not full equivalence',
        'Decimal JSON Val/CDbl: 16 residual-field mismatches per measured sample; not full equivalence',
        'JSON 30000-character chunks: failed Out of memory; not a timing result',
        'Accuracy_Checks intentionally changes mesh or requested checks; not an equivalence test',
        'Candidate01 and Candidate02 did not compile; only successful later candidates were measured',
    ],
    'archive': {'path': archive.name, 'bytes': archive.stat().st_size, 'sha256': sha(archive)},
    'files': manifest,
    'finalGeneratedModules': [
        {'name': p.name, 'sha256': sha(p)} for p in sorted((HERE / 'Modules').glob('*')) if p.is_file()
    ],
}
(HERE / 'EvidenceIndex.json').write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding='utf-8')
print(f'Archived {len(files)} evidence files; {archive.stat().st_size} bytes. All six reported plans complete; exceptions retained.')
