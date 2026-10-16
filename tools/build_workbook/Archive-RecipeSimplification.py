"""Архивирует доказательства перед удалением собственных временных копий книги."""
import hashlib
import argparse
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parents[2]
D = ROOT / 'docs/regression/ExcelGuardGuide_2026-10-13/RecipeSimplification_2026-10-16'


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--merge', action='store_true', help='Preserve every existing proof and add the final comment-only stage')
    args = parser.parse_args()
    assert json.loads((D/'Acceptance.json').read_text(encoding='utf-8-sig'))['PublishedOutputVerified']
    archive_path = D/'Evidence.zip'
    assert args.merge or not archive_path.exists(), 'Do not replace a verified archive.'
    previous = ZipFile(archive_path) if archive_path.exists() else None
    old_sha = hashlib.sha256(archive_path.read_bytes()).hexdigest() if previous else None
    if previous:
        old_receipt = json.loads((D/'ArchiveReceipt.json').read_text(encoding='utf-8-sig'))
        assert old_receipt['ArchiveSHA256'] == old_sha and previous.testzip() is None
    destination = D/'Evidence.next.zip' if previous else archive_path
    assert destination.parent.resolve() == D.resolve() and (not previous or not destination.exists())
    paths = sorted(path for path in D.rglob('*') if path.is_file() and path.suffix.lower() != '.xlsm'
                   and path.name not in ('Evidence.zip', 'Evidence.next.zip', 'ArchiveReceipt.json', 'ScratchCleanup.json', 'DeliveryReceipt.json'))
    current = {path.relative_to(D).as_posix(): path for path in paths}
    names = sorted(set(current) | (set(previous.namelist())-{'Manifest.json'} if previous else set()))
    manifest = {}
    with ZipFile(destination, 'w', compression=ZIP_DEFLATED, compresslevel=9) as archive:
        for name in names:
            content = current[name].read_bytes() if name in current else previous.read(name)
            if previous and name in previous.namelist() and previous.read(name) != content:
                old = previous.read(name)
                historical = 'History/BeforeGuideReadability/'+name
                assert historical not in names
                archive.writestr(historical, old)
                manifest[historical] = {'Bytes': len(old), 'SHA256': hashlib.sha256(old).hexdigest()}
            archive.writestr(name, content)
            manifest[name] = {'Bytes': len(content), 'SHA256': hashlib.sha256(content).hexdigest()}
        archive.writestr('Manifest.json', json.dumps(manifest, ensure_ascii=False, indent=2)+'\n')
    with ZipFile(destination) as archive:
        assert archive.testzip() is None
        for name, expected in manifest.items():
            actual = archive.read(name)
            assert len(actual) == expected['Bytes'] and hashlib.sha256(actual).hexdigest() == expected['SHA256']
    if previous:
        previous.close()
        destination.replace(archive_path)
    receipt = {'Status': 'ARCHIVE_VERIFIED', 'Entries': len(manifest), 'CRCVerified': True, 'EverySHA256Verified': True,
               'ArchiveSHA256': hashlib.sha256(archive_path.read_bytes()).hexdigest(), 'Bytes': archive_path.stat().st_size,
               'MergedPreviousSHA256': old_sha}
    (D/'ArchiveReceipt.json').write_text(json.dumps(receipt, indent=2)+'\n', encoding='utf-8')
    print(json.dumps(receipt))
