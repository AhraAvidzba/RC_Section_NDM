"""Сохраняет отчёты, отрицательные журналы и фактические контрольные книги изменения."""
import argparse
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parents[2]

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--directory', type=Path, required=True)
    args = parser.parse_args()
    directory = args.directory.resolve()
    allowed = (ROOT/'docs/regression/GeometryFull_2026-10-13').resolve()
    if not directory.is_relative_to(allowed):
        raise ValueError('Evidence is outside the authorized directory')
    target = ROOT/'docs/geometry_refactor/PlateauFactorEvidence_2026-10-13.zip'
    manifest_path = target.with_suffix('.json')
    if target.exists() or manifest_path.exists():
        raise ValueError('Existing evidence archive must not be replaced')
    paths = sorted(p for p in directory.rglob('*') if p.is_file() and p.suffix in ('.txt','.json','.log'))
    paths += [directory/'BaselineProbe/RC_Section_NDM.xlsm', directory/'ReleaseProbe/RC_Section_NDM.xlsm']
    paths += [ROOT/'src/Materials/CMaterialModelProvider.cls', ROOT/'tests/modTestMaterialDiagrams.bas',
              ROOT/'tests/modTestBatchCalculation.bas', ROOT/'tests/modTestWorkbookInterface.bas']
    paths += sorted((ROOT/'tools/build_workbook').glob('*PlateauFactor*'))
    paths += [ROOT/'tools/build_workbook/SettingsCatalog.ps1', ROOT/'tools/build_workbook/FocusedAcceptanceProbe.bas',
              ROOT/'tools/build_workbook/Prepare-FocusedAcceptance.ps1', ROOT/'docs/reports/PlateauFactor_2026-10-13.md']
    paths = sorted(set(paths))
    records = []
    with ZipFile(target, 'w', compression=ZIP_DEFLATED, compresslevel=9) as archive:
        for path in paths:
            data = path.read_bytes()
            name = path.relative_to(ROOT).as_posix()
            archive.writestr(name, data)
            records.append({'path':name, 'bytes':len(data), 'sha256':hashlib.sha256(data).hexdigest()})
    with ZipFile(target) as archive:
        if archive.testzip() is not None:
            raise ValueError('Archive CRC failed')
        for record in records:
            if hashlib.sha256(archive.read(record['path'])).hexdigest() != record['sha256']:
                raise ValueError('Archive SHA256 failed: '+record['path'])
    manifest = {'AllPassed':True, 'Archive':target.relative_to(ROOT).as_posix(),
                'ArchiveSHA256':hashlib.sha256(target.read_bytes()).hexdigest(), 'ArchiveBytes':target.stat().st_size,
                'Artifacts':records, 'CandidateWorkbook':'workbook/output/RC_Section_NDM.xlsm',
                'CandidateSHA256':hashlib.sha256((directory/'RC_Section_NDM.xlsm').read_bytes()).hexdigest(),
                'PreviousWorkbookSHA256':hashlib.sha256((directory/'Before.xlsm').read_bytes()).hexdigest(),
                'PreviousWorkbookRecovery':'Git 821ba077:workbook/output/RC_Section_NDM.xlsm',
                'Note':'Includes negative development logs; acceptance is given by NativeSourceEquality, UserDataV2, StatusComparison and Publication.'}
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps({'AllPassed':True,'Artifacts':len(records),'ArchiveBytes':manifest['ArchiveBytes']}, ensure_ascii=False))
