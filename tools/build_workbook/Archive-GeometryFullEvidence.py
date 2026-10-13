"""Archive complete acceptance evidence in small Git-compatible ZIP parts."""
import argparse
import hashlib
import json
import zipfile
from pathlib import Path


def sha(path):
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024*1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory', type=Path, required=True)
    parser.add_argument('--output-directory', type=Path, required=True)
    args = parser.parse_args()
    root, output = args.directory.resolve(), args.output_directory.resolve()
    if root == output or root in output.parents:
        raise ValueError('Archive output must be outside the evidence directory')
    if list(output.glob('GeometryFullEvidence_*.zip')) or (output/'GeometryFullArchive.json').exists():
        raise ValueError('Archives already exist; inspect the previous outcome before replacing evidence')
    output.mkdir(parents=True, exist_ok=True)
    artifacts, unique, archives = [], {}, []
    archive, archive_path, part_bytes, number = None, None, 0, 0
    budget = 64*1024*1024

    def finish():
        if archive is not None:
            archive.close()
            archives.append({'path': archive_path.name, 'bytes': archive_path.stat().st_size,
                             'sha256': sha(archive_path)})

    try:
        for path in sorted(root.rglob('*')):
            if not path.is_file() or '__pycache__' in path.parts:
                continue
            relative, size, digest = path.relative_to(root).as_posix(), path.stat().st_size, sha(path)
            if digest in unique:
                stored = unique[digest]
            else:
                if size > budget:
                    raise ValueError('Single artifact exceeds the archive part budget: ' + relative)
                if archive is None or part_bytes+size > budget:
                    finish()
                    number += 1
                    archive_path = output/f'GeometryFullEvidence_{number:02d}.zip'
                    archive = zipfile.ZipFile(archive_path, 'x', zipfile.ZIP_DEFLATED, compresslevel=6)
                    part_bytes = 0
                archive.write(path, relative)
                if sha(path) != digest:
                    raise ValueError('Evidence changed while archiving: ' + relative)
                part_bytes += size
                stored = {'archive': archive_path.name, 'entry': relative}
                unique[digest] = stored
            artifacts.append({'path': relative, 'bytes': size, 'sha256': digest, **stored})
        finish()
        archive = None
    finally:
        if archive is not None:
            archive.close()
    for item in archives:
        with zipfile.ZipFile(output/item['path']) as zipped:
            failure = zipped.testzip()
            if failure:
                raise ValueError('Corrupt archive entry: ' + failure)
    manifest = {'status': 'ARCHIVED_AND_CRC_VERIFIED', 'sourceDirectory': root.name,
                'artifacts': artifacts, 'archives': archives, 'uniqueBlobs': len(unique),
                'storage': 'Identical bytes stored once; use archive/entry for every original path',
                'restoration': 'Extract each archive, then copy its listed entry to each manifest path and verify SHA256'}
    (output/'GeometryFullArchive.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps({'status': manifest['status'], 'artifacts': len(artifacts),
                      'uniqueBlobs': len(unique), 'archives': len(archives),
                      'archiveBytes': sum(x['bytes'] for x in archives)}))


if __name__ == '__main__':
    main()
