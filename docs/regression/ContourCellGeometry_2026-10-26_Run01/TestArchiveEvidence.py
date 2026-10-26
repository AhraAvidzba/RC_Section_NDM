"""ДЛЯ ТЕСТОВ: упаковывает завершённые текстовые доказательства без изменения исходников.

Запускается root только после закрытия проверок. Сохраняет все исторические
результаты выбранных типов, включая FAIL; готовые книги и бинарные данные
не входят в компактный архив. Вложенные исторические Acceptance сохраняются.
"""
from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(__file__).resolve().parent
ZIP_PATH = ROOT / 'Logs.zip'
MANIFEST_PATH = ROOT / 'ArchiveManifest.json'
TEXT_SUFFIXES = {'.txt', '.log', '.json', '.ps1', '.py', '.bas', '.diff', '.md'}
EXCLUDED_DIRECTORIES = {'inputs', '__pycache__'}
FINAL_ROOT_STEMS = {'acceptance', 'processed', 'archivemanifest'}
CHUNK_BYTES = 1024 * 1024


def sha256_file(path: Path) -> str:
    """Считает хеш исходных байтов блоками, без декодирования и нормализации строк."""
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        for chunk in iter(lambda: stream.read(CHUNK_BYTES), b''):
            digest.update(chunk)
    return digest.hexdigest().upper()


def selected_files() -> list[Path]:
    """Выбирает только разрешённые файлы внутри Run01, не фильтруя результат теста."""
    result = []
    for directory, children, files in os.walk(ROOT, followlinks=False):
        children[:] = sorted(
            name for name in children
            if name.casefold() not in EXCLUDED_DIRECTORIES
            and not (Path(directory) / name).is_symlink()
        )
        for name in files:
            path = Path(directory) / name
            if path.suffix.casefold() not in TEXT_SUFFIXES:
                continue
            if path.name.casefold().endswith('_nativecode.json'):
                continue
            if path.parent == ROOT and path.stem.casefold() in FINAL_ROOT_STEMS:
                continue
            if path.is_symlink() or not path.resolve().is_relative_to(ROOT):
                raise RuntimeError('В доказательствах обнаружена ссылка за пределы Run01: ' + str(path))
            result.append(path)
    return sorted(result, key=lambda item: item.relative_to(ROOT).as_posix().casefold())


def write_archive(files: list[Path]) -> list[dict]:
    """Создаёт свежий ZIP и хеширует именно передаваемые в него байты каждого файла."""
    entries = []
    # Режим x не перезаписывает готовый или частично созданный предыдущий архив.
    with ZipFile(ZIP_PATH, mode='x', compression=ZIP_DEFLATED, compresslevel=9, allowZip64=True) as archive:
        for path in files:
            relative = path.relative_to(ROOT).as_posix()
            before = path.stat()
            digest = hashlib.sha256()
            count = 0
            with path.open('rb') as source, archive.open(relative, mode='w', force_zip64=True) as target:
                for chunk in iter(lambda: source.read(CHUNK_BYTES), b''):
                    target.write(chunk)
                    digest.update(chunk)
                    count += len(chunk)
            after = path.stat()
            if before.st_size != after.st_size or before.st_mtime_ns != after.st_mtime_ns or count != after.st_size:
                raise RuntimeError('Файл менялся во время архивации: ' + relative)
            entries.append({'Path': relative, 'Bytes': count, 'SHA256': digest.hexdigest().upper()})
    return entries


def verify_archive(entries: list[dict], original_files: list[Path]) -> None:
    """Повторно читает каждую запись ZIP и проверяет, что исходный набор не изменился."""
    expected = {entry['Path']: entry for entry in entries}
    with ZipFile(ZIP_PATH, mode='r') as archive:
        names = archive.namelist()
        if len(names) != len(expected) or set(names) != set(expected):
            raise RuntimeError('ZIP содержит пропущенные, лишние или повторные имена.')
        for name in names:
            digest = hashlib.sha256()
            count = 0
            with archive.open(name, mode='r') as source:
                for chunk in iter(lambda: source.read(CHUNK_BYTES), b''):
                    digest.update(chunk)
                    count += len(chunk)
            if count != expected[name]['Bytes'] or digest.hexdigest().upper() != expected[name]['SHA256']:
                raise RuntimeError('Не совпали размер/SHA256 записи ZIP: ' + name)
            if sha256_file(ROOT / name) != expected[name]['SHA256']:
                raise RuntimeError('Исходный файл изменился после упаковки: ' + name)
    if selected_files() != original_files:
        raise RuntimeError('Состав текстовых доказательств изменился во время архивации.')


def main() -> None:
    """Пишет отдельный манифест после успешной проверки свежего архива; PASS тестов не назначает."""
    if ROOT.name != 'ContourCellGeometry_2026-10-26_Run01':
        raise RuntimeError('Архиватор предназначен только для своего каталога Run01.')
    if ZIP_PATH.exists() or MANIFEST_PATH.exists():
        raise FileExistsError('Нужны новые Logs.zip и ArchiveManifest.json; существующие доказательства не перезаписываются.')
    files = selected_files()
    if not files:
        raise RuntimeError('Нет текстовых доказательств для упаковки.')
    entries = write_archive(files)
    verify_archive(entries, files)
    manifest = {
        'Status': 'ARCHIVED_TEXT_EVIDENCE_SHA256_VERIFIED',
        'ProducedBy': Path(__file__).name,
        'CreatedUTC': datetime.now(timezone.utc).isoformat(),
        'EvidenceDirectory': ROOT.name,
        'Archive': ZIP_PATH.name,
        'ArchiveBytes': ZIP_PATH.stat().st_size,
        'ArchiveSHA256': sha256_file(ZIP_PATH),
        'Compression': 'DEFLATE level 9',
        'EntryCount': len(entries),
        'UncompressedBytes': sum(entry['Bytes'] for entry in entries),
        'AllowedSuffixes': sorted(TEXT_SUFFIXES),
        'ExcludedDirectoriesAnyLevel': sorted(EXCLUDED_DIRECTORIES),
        'ExcludedFilenameSuffixes': ['_NativeCode.json'],
        'ExcludedRootFinalStems': sorted(FINAL_ROOT_STEMS),
        'HistoricalNestedAcceptanceIncluded': True,
        'ResultBasedFiltering': False,
        'SourceFilesChanged': False,
        'FullRegressionPassAssigned': False,
        'Entries': entries,
    }
    # Манифест исключён из ZIP: его хеш архива не создаёт рекурсивную зависимость.
    with MANIFEST_PATH.open('x', encoding='utf-8', newline='\n') as target:
        json.dump(manifest, target, ensure_ascii=False, indent=2)
        target.write('\n')
    print(json.dumps({key: manifest[key] for key in ('Status', 'Archive', 'EntryCount', 'ArchiveBytes', 'ArchiveSHA256')}))


if __name__ == '__main__':
    main()
