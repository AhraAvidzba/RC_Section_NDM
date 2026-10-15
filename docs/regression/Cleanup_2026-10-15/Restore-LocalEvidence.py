"""Восстанавливает один сохранённый протокол по его прежнему относительному адресу."""
import argparse
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
from zipfile import ZipFile

parser=argparse.ArgumentParser()
parser.add_argument('path',help='Прежний адрес относительно корня проекта, с /')
args=parser.parse_args()
root=Path.cwd().resolve(strict=True)
directory=root/'docs/regression/Cleanup_2026-10-15'
assert Path(__file__).resolve().parent == directory
parts=PurePosixPath(args.path).parts
assert not PurePosixPath(args.path).is_absolute() and '..' not in parts
assert args.path.startswith(('docs/regression/','tmp/'))
target=root.joinpath(*parts)
assert os.path.commonpath((str(root),str(target.resolve()))) == str(root)
assert not target.exists(), 'Existing files must not be overwritten'
for parent in target.parents:
    if parent==root:break
    if parent.exists():
        assert not parent.is_symlink() and not parent.stat(follow_symlinks=False).st_file_attributes & 0x400
with ZipFile(directory/'LocalDiagnostics.zip') as archive:
    record=next(item for item in json.loads(archive.read('manifest.json'))['Files'] if item['Path']==args.path)
    storage=record['Storage']
    if storage['Kind']=='LocalArchive':data=archive.read(storage['Member'])
    elif storage['Kind']=='TrackedFile':data=(root/storage['Path']).read_bytes()
    elif storage['Kind']=='TrackedArchive':
        with ZipFile(root/storage['Path']) as original:data=original.read(storage['Member'])
    else:raise ValueError('Unknown storage kind')
assert hashlib.sha256(data).hexdigest()==record['SHA256']
target.parent.mkdir(parents=True,exist_ok=True)
with target.open('xb') as stream:stream.write(data)
print('RESTORED: '+str(target))
