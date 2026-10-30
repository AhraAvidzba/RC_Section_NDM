"""ДЛЯ ТЕСТОВ: публикует компактные доказательства принятого Run03; книги, src и исторические FAIL не меняет."""
from pathlib import Path
from datetime import datetime, timezone
from zipfile import ZipFile, ZIP_DEFLATED
import hashlib
import json
import subprocess

ROOT = Path('C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM')
PRIVATE = Path(__file__).parent
LAUNCH = ROOT / 'docs/regression/FullSuiteLaunch_2026-10-29_UnifiedRebar_Run03'
FULL = ROOT / 'docs/regression/FullOnOff_2026-10-29_UnifiedRebar_Run03'
DEST = ROOT / 'docs/regression/UnifiedRebarPaths_2026-10-29_Run01/FinalRun03'


def TestJSON(path):
    """ДЛЯ ТЕСТОВ: читает полный исходный JSON без изменения доказательства."""
    return json.loads(path.read_text('utf-8-sig'))


def TestSHA(path):
    """ДЛЯ ТЕСТОВ: вычисляет хеш точных байтов файла."""
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def TestPrepare():
    """ДЛЯ ТЕСТОВ: допускает упаковку только после полного PASS, Ready30, 276 own и проверенного архива."""
    request = TestJSON(LAUNCH / 'LaunchRequest.json')
    complete = TestJSON(FULL / 'CompleteCurrentAcceptance.json')
    scope = TestJSON(LAUNCH / 'TerminalScopeReview.json')
    ready = TestJSON(PRIVATE / 'Ready30Publication03.json')
    own = TestJSON(PRIVATE / 'AllGoalOwnClosuresFinal01.json')
    archive = TestJSON(PRIVATE / 'Run03FinalArchive01/Manifest.json')
    assert request['SourceCommit'] == subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=ROOT, text=True).strip()
    assert complete['Status'] == 'PASS' and complete['SourceSHA256'] == request['SourceSHA256']
    assert complete['CoreSuites'] == 182 and complete['FocusedSingleCases'] == 122
    assert scope['PrimaryLogsReviewed'] and not scope['Failures'] and not scope['InventoryErrors']
    assert scope['CoreReportsRead'] == 12 and scope['FocusedRecordCount'] == 26
    assert own['AllOwnIdentitiesTerminal'] and own['UniqueOwnCount'] == 276
    assert ready['UserInputsResultsAndPlotPreserved'] and not ready['SolverRecalculated']
    assert TestSHA(Path(ready['Ready'])) == ready['ReadySHA256']
    assert TestSHA(ROOT / 'workbook/output/RC_Section_NDM.xlsm') == request['SourceSHA256']
    assert TestSHA(ROOT / 'workbook/output/VBA_All_Code.txt') == request['AllCodeSHA256']
    assert archive['Status'] == 'ENTIRE_ACCEPTED_RUN03_AND_READY30_PROOFS_ARCHIVED'
    assert TestSHA(Path(archive['Archive'])) == archive['ArchiveSHA256']
    assert not DEST.exists(), 'Inspect existing final evidence instead of replacing it'
    DEST.mkdir()
    copied = []

    def TestCopy(source, relative):
        """ДЛЯ ТЕСТОВ: копирует только явно перечисленный proof и проверяет равенство байтов."""
        assert source.is_file(), str(source)
        target = DEST / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        data = source.read_bytes(); target.write_bytes(data)
        assert target.read_bytes() == data
        copied.append({'Path': relative, 'OriginalPath': str(source), 'Bytes': len(data), 'SHA256': TestSHA(target)})

    for name in ('LaunchRequest.json', 'SourceEpoch.json', 'ReleaseEligibility.json', 'FullLaunch.json',
                 'Completion.json', 'DriverRun.json', 'HostRun.json', 'TerminalScopeReview.json',
                 'AllOwnClosuresReviewed.json', 'ScheduledTaskCleanup01.json'):
        TestCopy(LAUNCH / name, 'Launch/' + name)
    for name in ('Summary.json', 'CompleteCurrentAcceptance.json', 'Completion.json', 'Run.json',
                 'Focused/Run.json', 'Focused/Acceptance.json', 'Focused/CombinedAcceptance.json',
                 'Focused/Independent.json', 'Focused/Probe/Preparation.json', 'Focused/NativeCAD/Stage.json'):
        TestCopy(FULL / name, 'Acceptance/' + name)
    names = ('Ready30SchemaPreparation02.json', 'Ready30SchemaEquality02.json', 'Ready30FinalReview03.json',
             'Ready30Publication03.json', 'Ready30SchemaNative02/NativeVerification.json',
             'Ready30SchemaNative02/DropdownAcceptance.json', 'Ready30SchemaNative02/SchemaNativeAcceptance.json',
             'Ready30SchemaNative02/OwnClosure.json', 'Ready30SchemaNative02/Quit.json',
             'AllPreFullOwnClosures20.json', 'AllGoalOwnClosuresFinal01.json',
             'Run03TerminalStdIOReview01.json', 'Run03ObservationTransition01.json',
             'Ready30PreparationFailure01.json', 'Ready30NamedAnchorSelftest01.json',
             'Ready30MetadataCorrection01.json', 'FinalOwnGuardPreparation01.json',
             'Run03TaskBeforeCleanup01.xml', 'Run03FinalArchive01/Manifest.json')
    for name in names:
        TestCopy(PRIVATE / name, 'Proofs/' + name)
    own_recheck = Path(ready['OwnRecheckProof']); assert own_recheck.is_relative_to(PRIVATE)
    TestCopy(own_recheck, 'Proofs/' + own_recheck.name)
    for name in ('TestPrepareReady30Schema02.py', 'TestReady30SchemaTools02.py', 'TestPrepareReady30Current01.py',
                 'TestReady30MainSchemaSelftest03.py', 'TestReady30NamedAnchorSelftest01.py',
                 'TestNativeReady30Schema02.ps1', 'TestVerifySource.ps1', 'TestReviewReady30Final03.py',
                 'TestPublishReady30Final03.ps1', 'TestObserveOwnIdentityFinal01.ps1',
                 'TestReviewAllGoalOwnFinal01.ps1', 'TestCleanupFullRun03Task01.ps1',
                 'TestArchiveRun03Final01.py', 'TestPrepareFinalRun03Evidence01.py'):
        TestCopy(PRIVATE / name, 'Tools/' + name)

    files = {ROOT / row['Path']: 'Repository/' + row['Path'] for row in scope['Primary']}
    for row in scope['Primary']:
        assert row['AllBytesRead'] and TestSHA(ROOT / row['Path']) == row['SHA256']
    for name in ('FullLaunch.stdout.log', 'FullLaunch.stderr.log'):
        files[LAUNCH / name] = 'Repository/' + (LAUNCH / name).relative_to(ROOT).as_posix()
    files[PRIVATE / 'Ready30SchemaNative02/VBA_All_Code.txt'] = 'Ready30/VBA_All_Code.txt'
    package_path = DEST / 'ReviewedPrimaryAndReadyCode.zip'; members = []
    with ZipFile(package_path, 'x', compression=ZIP_DEFLATED, compresslevel=6) as package:
        for source, name in sorted(files.items(), key=lambda item: item[1]):
            data = source.read_bytes(); package.writestr(name, data)
            members.append({'Path': name, 'OriginalPath': str(source), 'Bytes': len(data), 'SHA256': TestSHA(source)})
    with ZipFile(package_path) as package:
        assert package.testzip() is None
        assert package.namelist() == [row['Path'] for row in members]
        for row in members:
            assert hashlib.sha256(package.read(row['Path'])).hexdigest().upper() == row['SHA256']
            assert TestSHA(Path(row['OriginalPath'])) == row['SHA256']
    manifest = {'Status': 'EXACT_PRIMARY_AND_READY_CODE_ARCHIVE_VERIFIED', 'Files': members,
                'ArchiveSHA256': TestSHA(package_path), 'ArchiveBytes': package_path.stat().st_size}
    (DEST / 'ReviewedPrimaryAndReadyCode.Manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    completion = TestJSON(LAUNCH / 'Completion.json')
    start = datetime.fromisoformat(request['Started']); end = datetime.fromisoformat(completion['FinishedUTC'])
    result = {'Status': 'PASS_ALL_REQUIRED_ACCEPTANCE_AND_READY30_PUBLISHED',
              'ObservedUTC': datetime.now(timezone.utc).isoformat(), 'AcceptedSourceCommit': request['SourceCommit'],
              'MainSHA256': request['SourceSHA256'], 'AllCodeSHA256': request['AllCodeSHA256'],
              'Ready30SHA256': ready['ReadySHA256'], 'CoreGroups': 12, 'CoreMacros': 182,
              'FocusedRecords': 26, 'FocusedSingleCases': 122, 'AllPrimaryFilesRead': len(scope['Primary']),
              'FullRunMinutes': (end - start).total_seconds() / 60,
              'CorePassedAssertions': complete['CorePassedAssertions'],
              'FocusedPassedAssertions': complete['FocusedPassedAssertions'],
              'IndependentIntegrals': 560, 'IndependentModels': 5, 'ExpectedNumFailStatesIndividuallyAccepted': 10,
              'IndependentUnsupported': 0, 'FullOwnTerminal': 44, 'AllGoalOwnTerminal': 276,
              'Original505InputsUnchanged': True, 'ReadyDataResultsAndPlotPreserved': True,
              'Ready30OriginalLoadCount': 30, 'Ready30HistoricalInputErrIDs': ['PERF_14','PERF_21','PERF_27','PERF_28'],
              'ReadySolverRecalculated': False, 'MainClean': True, 'OptionalNDMCompanionPublished': False,
              'HistoricalRun02StillReviewRequired': True, 'PrivateFullArchiveSHA256': archive['ArchiveSHA256'],
              'PrimaryArchiveSHA256': manifest['ArchiveSHA256'], 'PrimaryArchiveBytes': manifest['ArchiveBytes'],
              'UserMainBackupPrivatePreserved': True, 'ReadyOriginalBackupPrivatePreserved': True,
              'FeatureBranch': 'codex/unified-rebar-paths', 'FeaturePushVerified': False,
              'MainIntegrationRequiresUserAgreement': True, 'CopiedProofs': copied}
    (DEST / 'Processed.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    (DEST / 'README.md').write_text('Доказательства полного Run03 на D71 и обновления Ready30 на 2CD.\n\n'
        'ReviewedPrimaryAndReadyCode.zip содержит все 40 прочитанных первичных файлов, stdout/stderr и нативный экспорт Ready30; Manifest проверяет точные байты. '
        'Полный архив 383 файлов сохранён приватно, его хеш и состав находятся в Proofs/Run03FinalArchive01/Manifest.json.\n\n'
        'Tools — зафиксированные одноразовые инструменты и их зависимости; они привязаны к исходным приватным путям и SHA. '
        'Не запускать их из каталога доказательств. Исторические FAIL и резервные пользовательские книги сохранены отдельно.\n', encoding='utf-8')
    print(json.dumps({key: result[key] for key in ('Status','FullRunMinutes','AllPrimaryFilesRead','AllGoalOwnTerminal','PrimaryArchiveBytes','PrimaryArchiveSHA256')}))


if __name__ == '__main__':
    TestPrepare()
