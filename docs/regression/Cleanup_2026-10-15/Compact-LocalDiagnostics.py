"""Сохраняет уникальные локальные протоколы в проверенный архив и убирает распакованные копии."""
import collections
import hashlib
import json
import os
import subprocess
import time
import zlib
from pathlib import Path, PurePosixPath
from zipfile import ZipFile, ZIP_DEFLATED

root = Path.cwd().resolve(strict=True)
directory = root/'docs/regression/Cleanup_2026-10-15'
assert Path(__file__).resolve().parent == directory
inventory = json.loads((root/'tmp/Cleanup_2026-10-15/Inventory.json').read_text(encoding='utf-8'))
tracked = set(subprocess.check_output(['git','ls-files','-z'],cwd=root).decode('utf-8').split('\0'))
original_status = subprocess.check_output(['git','status','--porcelain'],cwd=root)
critical = json.loads((directory/'TemporaryBooks.json').read_text(encoding='utf-8'))['ProtectedHashes']

def digest(path):
    value = hashlib.sha256()
    crc = 0
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1024*1024), b''):
            value.update(block)
            crc = zlib.crc32(block,crc)
    return value.hexdigest(), crc & 0xffffffff

def safe_path(relative):
    parts = PurePosixPath(relative).parts
    assert not PurePosixPath(relative).is_absolute() and '..' not in parts
    assert parts[0] not in ('.git','norms','reference','control_examples')
    path = root.joinpath(*parts)
    assert os.path.commonpath((str(root),str(path.resolve(strict=True)))) == str(root)
    for item in (path,*path.parents):
        if item == root: break
        assert not item.is_symlink() and not item.stat(follow_symlinks=False).st_file_attributes & 0x400
    return path

candidates = []
cache = []
for item in inventory['Files']:
    relative = item['Path']
    if relative in tracked or relative.startswith('.git/'): continue
    if relative.startswith('tmp/Cleanup_2026-10-15/'): continue
    if '/vendor/' in relative: continue
    if not (root/relative).exists(): continue
    if '__pycache__' in PurePosixPath(relative).parts or relative.endswith(('.pyc','.pyo')):
        cache.append(item)
    elif relative.startswith(('docs/regression/','tmp/')):
        assert not relative.startswith('docs/regression/Cleanup_2026-10-15/')
        assert Path(relative).suffix.lower() != '.xlsm'
        candidates.append(item)

paths = candidates + cache
ignored = set(subprocess.run(['git','check-ignore','--stdin','-z'],cwd=root,
    input=('\0'.join(item['Path'] for item in paths)+'\0').encode('utf-8'),
    stdout=subprocess.PIPE,check=True).stdout.decode('utf-8').split('\0'))
for item in paths:
    assert item['Path'] in ignored
    path = safe_path(item['Path'])
    stat = path.stat()
    assert stat.st_size == item['Bytes'] and stat.st_mtime_ns == item['ModifiedNs'],item['Path']

# Существующие отслеживаемые файлы/архивы остаются по своим адресам.
by_size = collections.defaultdict(list)
archive_entries = collections.defaultdict(list)
for item in inventory['Files']:
    if not item['Tracked']: continue
    by_size[item['Bytes']].append(item['Path'])
    if item['Path'].endswith('.zip'):
        with ZipFile(root/item['Path']) as archive:
            for member in archive.infolist():
                if not member.is_dir():
                    archive_entries[(member.file_size,member.CRC)].append((item['Path'],member.filename))

tracked_hashes = {}
member_hashes = {}
unique = {}
manifest = []
read_bytes = 0
last_print = time.monotonic()
archive_path = directory/'LocalDiagnostics.zip'
assert not archive_path.exists(), 'Existing archive must not be overwritten'
print('ARCHIVE_START: files=%d; bytes=%d' % (len(candidates),sum(x['Bytes'] for x in candidates)),flush=True)
with ZipFile(archive_path,'x',ZIP_DEFLATED,compresslevel=1,allowZip64=True) as target:
    for index,item in enumerate(candidates,1):
        relative = item['Path']
        path = safe_path(relative)
        checksum,crc = digest(path)
        record = dict(item, SHA256=checksum)
        storage = unique.get(checksum)
        if storage is None:
            for existing in by_size[item['Bytes']]:
                if existing not in tracked_hashes: tracked_hashes[existing] = digest(root/existing)[0]
                if checksum == tracked_hashes[existing]:
                    storage = {'Kind':'TrackedFile','Path':existing}
                    break
        if storage is None:
            for existing,member in archive_entries[(item['Bytes'],crc)]:
                key = (existing,member)
                if key not in member_hashes:
                    with ZipFile(root/existing) as archive, archive.open(member) as stream:
                        member_hashes[key] = hashlib.file_digest(stream,'sha256').hexdigest()
                if checksum == member_hashes[key]:
                    storage = {'Kind':'TrackedArchive','Path':existing,'Member':member}
                    break
        if storage is None:
            member = 'files/'+relative
            target.write(path,member)
            storage = {'Kind':'LocalArchive','Member':member}
        unique[checksum] = storage
        record['Storage'] = storage
        manifest.append(record)
        read_bytes += item['Bytes']
        if time.monotonic()-last_print >= 15:
            print('ARCHIVE_PROGRESS: files=%d/%d; readGiB=%.3f' % (index,len(candidates),read_bytes/2**30),flush=True)
            last_print=time.monotonic()
    target.writestr('manifest.json',json.dumps({'Version':1,'Files':manifest},ensure_ascii=False,indent=2))

print('ARCHIVE_VERIFY: unique=%d; archiveMiB=%.2f' % (len(unique),archive_path.stat().st_size/2**20),flush=True)
verified = set()
with ZipFile(archive_path) as archive:
    restored_manifest=json.loads(archive.read('manifest.json'))['Files']
    assert restored_manifest == manifest
    for record in manifest:
        storage=record['Storage']
        if storage['Kind'] != 'LocalArchive' or storage['Member'] in verified: continue
        with archive.open(storage['Member']) as stream:
            assert hashlib.file_digest(stream,'sha256').hexdigest() == record['SHA256']
        verified.add(storage['Member'])

deleted = []
skipped = []
for index,record in enumerate(manifest,1):
    relative=record['Path']; path=safe_path(relative); stat=path.stat()
    if stat.st_size != record['Bytes'] or stat.st_mtime_ns != record['ModifiedNs'] or digest(path)[0] != record['SHA256']:
        skipped.append({'Path':relative,'Reason':'Changed during archival'}); continue
    try:
        path.unlink(); deleted.append(relative)
    except OSError as exc: skipped.append({'Path':relative,'Reason':str(exc)})
for item in cache:
    path=safe_path(item['Path']); stat=path.stat()
    if stat.st_size != item['Bytes'] or stat.st_mtime_ns != item['ModifiedNs']:
        skipped.append({'Path':item['Path'],'Reason':'Cache changed'}); continue
    try:
        path.unlink(); deleted.append(item['Path'])
    except OSError as exc: skipped.append({'Path':item['Path'],'Reason':str(exc)})
deleted_set=set(deleted)
freed=sum(item['Bytes'] for item in paths if item['Path'] in deleted_set)
assert all(digest(root/path)[0] == checksum for path,checksum in critical.items())
assert subprocess.check_output(['git','status','--porcelain'],cwd=root) == original_status
result={'Passed':not skipped,'ArchivedFiles':len(manifest),'UniqueContents':len(unique),
        'StoredPayloads':len(verified),'ArchiveBytes':archive_path.stat().st_size,
        'RemovedFiles':len(deleted),'RemovedBytes':freed,'NetFreedBytes':freed-archive_path.stat().st_size,
        'SourceCodeBookAllCodeAndGitStatusUnchanged':True,'Skipped':skipped,
        'ArchiveSHA256':digest(archive_path)[0]}
(directory/'Diagnostics.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(result,ensure_ascii=False),flush=True)
