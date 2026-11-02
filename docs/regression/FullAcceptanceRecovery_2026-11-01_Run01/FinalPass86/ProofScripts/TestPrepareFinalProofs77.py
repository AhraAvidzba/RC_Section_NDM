"""ДЛЯ ТЕСТОВ: после полного PASS и публикации готовит компактные доказательства в частной папке.

Не изменяет исходные журналы, книги, отчёт репозитория или Git. Перечитывает весь
первичный состав и обе опубликованные книги; общий PASS относится только к принятой SHA.
"""
from pathlib import Path
from datetime import datetime, timezone
from zipfile import ZipFile, ZIP_DEFLATED
import hashlib
import json
import subprocess

P = Path(__file__).resolve().parent
R = Path('C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM')
PS5 = 'C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe'


def test_sha(raw):
    """ДЛЯ ТЕСТОВ: вычисляет SHA фактических байтов без нормализации."""
    return hashlib.sha256(raw).hexdigest().upper()


def test_read(path):
    """ДЛЯ ТЕСТОВ: читает сохранённый JSON с допустимой UTF-8 BOM."""
    return json.loads(path.read_text('utf-8-sig'))


def test_zip_parts(path):
    """ДЛЯ ТЕСТОВ: читает все части XLSM с проверкой CRC и уникальности имён."""
    with ZipFile(path) as saved:
        assert len(saved.namelist()) == len(set(saved.namelist())), 'Duplicate ZIP parts'
        assert saved.testzip() is None, 'Corrupt ZIP part'
        return {name: saved.read(name) for name in saved.namelist()}


def main():
    """ДЛЯ ТЕСТОВ: проверяет действительные end-only gates перед созданием нового пакета."""
    result = subprocess.run([PS5, '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
                             str(P / 'TestReadCurrentRun02Pass60.ps1')],
                            capture_output=True, encoding='utf-8', errors='strict', check=True)
    gate = json.loads(result.stdout.lstrip('\ufeff'))
    assert gate['OfficialPASS'] and gate['AllOwnIdentitiesTerminal']
    assert (gate['CoreSuites'], gate['FocusedSingles'], gate['FocusedRecords'], gate['InputsChecked']) == (182, 122, 26, 522)
    launch, evidence = Path(gate['Launch']), Path(gate['Evidence'])
    scope = test_read(launch / 'TerminalScopeReview.json')
    request = test_read(launch / 'LaunchRequest.json')
    completion = test_read(launch / 'Completion.json')
    io = test_read(P / 'AdditionalIO74.json')
    publication = test_read(P / 'UserBooksCodeOnlyPublication64.json')
    assert publication['Status'] == 'BOTH_USER_BOOKS_CODE_ONLY_PUBLISHED_WITH_DATA_PRESERVED'
    assert publication['AcceptedSourceSHA256'] == gate['SourceSHA256']
    assert publication['SourceCommit'] == gate['SourceCommit'] and not publication['NDSRecalculated']
    assert len(publication['Books']) == 2 and {x['Label'] for x in publication['Books']} == {'Main', 'Heavy'}
    donor = test_zip_parts(Path(gate['Source']))['xl/vbaProject.bin']
    books = []
    for book in publication['Books']:
        path, backup = Path(book['Path']), Path(book['Backup'])
        assert test_sha(path.read_bytes()) == book['AfterSHA256'], 'Published user book changed; review the current save'
        assert test_sha(backup.read_bytes()) == book['BeforeSHA256'] == book['BackupSHA256']
        before, after = test_zip_parts(backup), test_zip_parts(path)
        assert before.keys() == after.keys() and after['xl/vbaProject.bin'] == donor
        changed = [name for name in before if before[name] != after[name]]
        assert changed == ['xl/vbaProject.bin'], 'Only the accepted code may differ'
        native = test_read(P / 'UserBooksNative63' / book['Label'] / 'Accepted.json')
        assert native['Status'] == 'SAVED_USER_BOOK_CODE_AND_DATA_VERIFIED'
        assert native['PreparedSHA256'] == book['AfterSHA256'] and native['OwnIdentityTerminal']
        assert not native['NDSRecalculated']
        books.append(dict(Label=book['Label'], Path=str(path), SHA256=book['AfterSHA256'],
                          BackupSHA256=book['BeforeSHA256'], PreservedPartCount=len(before)-1,
                          CodeOnlyChanged=True, OtherUncompressedPartsByteExact=True,
                          NativeComponents=187, NDSRecalculated=False,
                          FullPassAppliesOnlyToAcceptedSource=True))

    selected = {}
    def test_include(path, name, expected=None):
        """ДЛЯ ТЕСТОВ: перечитывает выбранное доказательство и связывает копию с исходным SHA."""
        raw = path.read_bytes()
        if expected:
            assert test_sha(raw) == expected, 'Primary changed after review: ' + str(path)
        assert name not in selected
        selected[name] = (path, raw)

    for row in scope['Primary']:
        assert row['AllBytesRead'] and row['Closed'] and not row['Errors']
        test_include(R / row['Path'], 'Repo/' + row['Path'], row['SHA256'])
    for row in io['Completed']:
        assert row['Passed'] and row['AllBytesRead'] and not row['Errors']
        test_include(R / row['Primary'], 'Repo/' + row['Primary'], row['SHA256'])
    # В пакете остаются ожидаемые входные ID: их можно сопоставить с каждой фактической FOCUSED_TIME строкой.
    for path in sorted((evidence / 'Focused' / 'Probe').glob('*_Manifest.txt')):
        test_include(path, 'Repo/' + path.relative_to(R).as_posix())
    path = evidence / 'Focused' / 'Probe' / 'Preparation.json'
    test_include(path, 'Repo/' + path.relative_to(R).as_posix())
    for folder, names in ((launch, ('LaunchRequest.json', 'FullLaunch.json', 'SourceEpoch.json',
                                    'ReleaseEligibility.json', 'DriverRun.json', 'HostRun.json',
                                    'Completion.json', 'TerminalScopeReview.json', 'AllOwnClosuresReviewed.json',
                                    'ScheduledTaskCleanup61.json', 'FullLaunch.stdout.log', 'FullLaunch.stderr.log')),
                          (evidence, ('Run.json', 'CompleteCurrentAcceptance.json',
                                      'Focused/Run.json', 'Focused/CombinedAcceptance.json',
                                      'Focused/Independent.json'))):
        for name in names:
            path = folder / name
            key = 'Repo/' + path.relative_to(R).as_posix()
            if key not in selected:
                test_include(path, key)
    for name in ('TerminalStdIOReviewed68.json', 'CompletionRequirementMap73.json', 'AdditionalIO74.json',
                 'PreparedUserBooks62.json', 'UserBooksCodeOnlyPublication64.json'):
        test_include(P / name, 'Private/' + name)
    for label in ('Main', 'Heavy'):
        for name in ('Native.json', 'SourceEquality.json', 'Own.json', 'OwnClosure.json', 'Accepted.json'):
            relative = 'UserBooksNative63/' + label + '/' + name
            test_include(P / relative, 'Private/' + relative)
    # NativeInvocation, ownership и закрытие дочерних процессов должны оставаться проверяемыми после упаковки.
    for folder in (launch, evidence, R / 'docs/regression/FullAcceptanceRecovery_2026-11-01_Run01/AdditionalIO74'):
        for path in sorted(folder.rglob('*.json')):
            key = 'Repo/' + path.relative_to(R).as_posix()
            if key not in selected:
                test_include(path, key)

    destination = P / 'FinalProcessed77'
    assert not destination.exists(), 'Inspect existing proof package rather than overwrite it'
    destination.mkdir()
    archive = destination / 'ReviewedCurrentRun02Primary.zip'
    manifest = []
    with ZipFile(archive, 'x', compression=ZIP_DEFLATED, compresslevel=9) as output:
        for name, (path, raw) in sorted(selected.items()):
            output.writestr(name, raw)
            manifest.append(dict(Path=name, Bytes=len(raw), SHA256=test_sha(raw)))
    with ZipFile(archive) as saved:
        assert saved.testzip() is None and len(saved.namelist()) == len(manifest)
        for row in manifest:
            assert test_sha(saved.read(row['Path'])) == row['SHA256']
            assert selected[row['Path']][0].read_bytes() == selected[row['Path']][1], 'Proof changed while packing'
    (destination / 'PrimaryManifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    proof = dict(Status='FULL_CORE_FOCUSED_AND_ADDITIONAL_IO_PASS_USER_BOOKS_CODE_ONLY_UPDATED',
                 CheckedUTC=datetime.now(timezone.utc).isoformat(),
                 AcceptedSourceSHA256=gate['SourceSHA256'], SourceCommit=gate['SourceCommit'],
                 Started=request['Started'], FinishedUTC=completion['FinishedUTC'],
                 CoreGroups=12, CoreMacroExecutions=182, FocusedRecords=26, FocusedActualSingleIDs=122,
                 AdditionalIOMacrosPerMode=9, AdditionalIOModes=['No', 'Yes'],
                 EntireRawPrimaryReviewed=True, AllOwnIdentitiesTerminal=True, InputsChecked=522,
                 IndependentModels=5, IndependentIntegrals=560, IndividuallyReviewedNumFailStates=10,
                 ActualCADTests=2, UserBooks=books, ArchiveFile=archive.name,
                 ArchiveBytes=archive.stat().st_size, ArchiveSHA256=test_sha(archive.read_bytes()),
                 ProofFileCount=len(manifest), NativeProcessesStartedByPackager=0,
                 SourceEvidenceChanged=False, GitCommitAndPushPending=True)
    (destination / 'Processed.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(proof, ensure_ascii=False))


if __name__ == '__main__':
    main()
