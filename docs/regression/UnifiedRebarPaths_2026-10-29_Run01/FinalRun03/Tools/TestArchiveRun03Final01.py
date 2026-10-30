"""ДЛЯ ТЕСТОВ: архивирует принятый Run03 и конечные proofs, не меняя исходные журналы."""
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import hashlib
import json
import subprocess

ROOT = Path('C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM')
PRIVATE = Path(__file__).parent
LAUNCH = ROOT / 'docs/regression/FullSuiteLaunch_2026-10-29_UnifiedRebar_Run03'
COMMIT = '8910245966806984e349b004cc8e0da0d4dccb86'
SOURCE = 'D71B27FDA6A7EAF1F8EC708D07111524171C14ACDF1C711B6278491D64E91FBE'


def TestJSON(path):
    """ДЛЯ ТЕСТОВ: читает весь JSON с допустимым UTF8 BOM."""
    return json.loads(path.read_text('utf-8-sig'))


def TestSHA(data):
    """ДЛЯ ТЕСТОВ: считает хеш точных байтов доказательства."""
    return hashlib.sha256(data).hexdigest().upper()


def TestArchive():
    """ДЛЯ ТЕСТОВ: требует полный PASS, Ready30 и точный закрытый epoch до упаковки."""
    request = TestJSON(LAUNCH / 'LaunchRequest.json')
    evidence = (ROOT / request['EvidenceDirectory']).resolve()
    assert evidence == ROOT / 'docs/regression/FullOnOff_2026-10-29_UnifiedRebar_Run03'
    complete = TestJSON(evidence / 'CompleteCurrentAcceptance.json')
    completion = TestJSON(LAUNCH / 'Completion.json')
    scope = TestJSON(LAUNCH / 'TerminalScopeReview.json')
    own = TestJSON(LAUNCH / 'AllOwnClosuresReviewed.json')
    goal_own = TestJSON(PRIVATE / 'AllGoalOwnClosuresFinal01.json')
    ready = TestJSON(PRIVATE / 'Ready30Publication03.json')
    cleanup = TestJSON(LAUNCH / 'ScheduledTaskCleanup01.json')
    assert request['SourceCommit'] == COMMIT and request['SourceSHA256'] == SOURCE
    assert complete['Status'] == 'PASS' and complete['SourceSHA256'] == SOURCE
    assert complete['CoreSuites'] == 182 and complete['FocusedSingleCases'] == 122
    assert completion['FinishedUTC'] and completion['SourceSHA256'] == SOURCE
    assert completion['CompleteAcceptancePassed'] and not completion['ChangedInputs']
    assert not completion['HeadChanged'] and not completion['DriverError']
    assert all(completion[key] == 0 for key in ('RunnerExit', 'SummaryExit', 'FinalizeExit'))
    assert scope['PrimaryLogsReviewed'] and scope['CoreReportsRead'] == 12
    assert scope['CoreMacrosReviewed'] == 182 and scope['FocusedRecordCount'] == 26
    assert scope['FocusedSingleCasesReviewed'] == 122
    assert not scope['Failures'] and not scope['InventoryErrors']
    independent = TestJSON(evidence / 'Focused/Independent.json')
    assert independent['Passed'] and independent['ModelCount'] == 5 and not independent['Unsupported']
    assert independent['ReviewedNumFailStates'] == 10 and len(independent['Failures']) == 10
    assert all(row['Passed'] for row in independent['Failures'])
    assert own['AllOwnIdentitiesTerminal'] and own['InputsUnchanged'] and own['InputsChecked'] == 505
    assert own['ScopeReviewSHA256'] == TestSHA((LAUNCH / 'TerminalScopeReview.json').read_bytes())
    assert goal_own['AllOwnIdentitiesTerminal'] and not goal_own['UnknownAcceptedAsAbsent']
    assert not goal_own['AnyProcessClosed'] and goal_own['SourceSHA256'] == SOURCE
    assert ready['Status'] == 'PUBLISHED_CURRENT_CODE_READY30_WITH_ORIGINAL_DATA'
    assert ready['AcceptedMainSHA256'] == SOURCE and ready['UserInputsResultsAndPlotPreserved']
    assert not ready['SolverRecalculated'] and ready['NativeOwnTerminal']
    assert TestSHA(Path(ready['Ready']).read_bytes()) == ready['ReadySHA256']
    assert cleanup['Status'] in ('EXACT_CURRENT_FULL_TASK_REMOVED_AFTER_ALL_OWN_EXIT',
                                 'ALREADY_ABSENT_AFTER_CURRENT_OWN_CLOSURES')
    assert cleanup['TaskName'] == 'RC_NDM_Full_20261029_UnifiedRebar_Run03'
    head = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=ROOT, text=True).strip()
    assert head == COMMIT, 'Seal original epoch before the final documentation commit'
    inputs = request['ProtectedInputs'] + request['SourceComponents'] + request['ParticipatingScripts']
    assert len(inputs) == 505
    for row in inputs:
        assert TestSHA((ROOT / row['Path']).read_bytes()) == row['SHA256'], row['Path']
    for row in scope['Primary']:
        assert row['AllBytesRead'] and TestSHA((ROOT / row['Path']).read_bytes()) == row['SHA256']

    selected = [
        'AllPreFullOwnClosures20.json', 'AllGoalOwnClosuresFinal01.json',
        'Ready30SchemaPreparation02.json', 'Ready30SchemaEquality02.json',
        'Ready30FinalReview03.json', 'Ready30Publication03.json',
        'Ready30SchemaNative02/NativeVerification.json',
        'Ready30SchemaNative02/OwnClosure.json', 'Ready30SchemaNative02/Quit.json',
        'Ready30SchemaNative02/VBA_All_Code.txt', 'Run03TaskBeforeCleanup01.xml',
        'FinalOwnGuardPreparation01.json', 'TestObserveOwnIdentityFinal01.ps1',
        'TestReviewAllGoalOwnFinal01.ps1', 'TestPublishReady30Final03.ps1',
        'TestArchiveRun03Final01.py', 'Run03TerminalStdIOReview01.json',
        'Run03ObservationTransition01.json', 'Ready30PreparationFailure01.json',
        'Ready30NamedAnchorSelftest01.json', 'Ready30MetadataCorrection01.json',
        'TestReady30NamedAnchorSelftest01.py', 'TestPrepareReady30Schema02.py',
        'TestReady30SchemaTools02.py', 'TestReviewReady30Final03.py',
        'TestNativeReady30Schema02.ps1',
    ]
    own_recheck = Path(ready['OwnRecheckProof']).resolve()
    assert own_recheck.is_relative_to(PRIVATE.resolve())
    assert TestSHA(own_recheck.read_bytes()) == ready['OwnRecheckSHA256']
    paths = {path: 'Repository/' + path.relative_to(ROOT).as_posix()
             for folder in (LAUNCH, evidence) for path in folder.rglob('*') if path.is_file()}
    for relative in selected:
        path = PRIVATE / relative
        assert path.is_file(), relative
        paths[path] = 'PrivateProofs/' + relative
    paths[own_recheck] = 'PrivateProofs/' + own_recheck.name
    target = PRIVATE / 'Run03FinalArchive01'
    assert not target.exists(), 'Inspect an existing archive receipt instead of replacing it'
    target.mkdir()
    archive = target / 'ReviewedFullRun03AndReady30.zip'
    rows = []
    with ZipFile(archive, 'x', compression=ZIP_DEFLATED, compresslevel=6) as package:
        for path, name in sorted(paths.items(), key=lambda item: item[1]):
            data = path.read_bytes()
            package.writestr(name, data)
            rows.append({'Path': name, 'OriginalPath': str(path), 'Bytes': len(data), 'SHA256': TestSHA(data)})
    with ZipFile(archive) as package:
        assert package.testzip() is None
        assert package.namelist() == [row['Path'] for row in rows]
        for row in rows:
            assert TestSHA(package.read(row['Path'])) == row['SHA256']
            assert TestSHA(Path(row['OriginalPath']).read_bytes()) == row['SHA256']
    result = {
        'Status': 'ENTIRE_ACCEPTED_RUN03_AND_READY30_PROOFS_ARCHIVED',
        'ObservedUTC': datetime.now(timezone.utc).isoformat(),
        'SourceSHA256': SOURCE, 'SourceCommit': COMMIT,
        'CoreMacrosRead': 182, 'FocusedSingleCasesRead': 122, 'FocusedRecordsRead': 26,
        'OriginalCompletionStatus': completion['Status'], 'AcceptanceStatus': complete['Status'],
        'ReadySHA256': ready['ReadySHA256'], 'AllOwnIdentitiesTerminal': True,
        'OriginalEvidenceRewritten': False, 'Archive': str(archive),
        'ArchiveBytes': archive.stat().st_size, 'ArchiveSHA256': TestSHA(archive.read_bytes()),
        'FileCount': len(rows), 'Files': rows,
    }
    (target / 'Manifest.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({key: result[key] for key in ('Status', 'FileCount', 'ArchiveBytes', 'ArchiveSHA256')}))


if __name__ == '__main__':
    TestArchive()
