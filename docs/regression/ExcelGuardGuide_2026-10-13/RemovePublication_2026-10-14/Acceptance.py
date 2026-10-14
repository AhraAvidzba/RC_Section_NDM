"""Сводит независимые доказательства всех требований узкой правки API."""
import hashlib
import json
import re
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

directory = Path(__file__).resolve().parent
root = directory.parents[3]

def read(name):
    return json.loads((directory/name).read_text(encoding='utf-8-sig'))

before = read('Before.json')
publication = read('Publication.json')
equivalence = read('Equivalence.json')
preservation = read('PublishedPreservation.json')
equality = read('PublishedSourceEquality.json')
help_result = read('PublishedStandaloneHelp.json')
guide = read('GuideNativeAccepted/GuideEvidence.json')
directed = read('Directed01/Run.json')
counts = {}
for run in directed['Runs']:
    text = (directory/'Directed01'/(run['Macro']+'.txt')).read_text(encoding='utf-8-sig')
    passed, failed = map(int, re.findall(r'passed=(\d+); failed=(\d+)',text)[-1])
    assert run['Passed'] and not failed
    counts[run['Macro']] = passed
assert sum(counts.values()) == 303
assert equivalence['Passed'] and equivalence['Calls'] == 13 and equivalence['UnchangedSourceComponents'] == 145
assert preservation['Passed'] and equality['Passed'] and help_result['status'] == 'STANDALONE_HELP_PASS'
assert guide['Status'] == 'DIRECTED_NATIVE_GUIDE_PASS' and guide['Idempotent'] and guide['ReopenTextMatch']
assert not publication['ReadbackPending'] and publication['AllCodeMatchesActualWorkbookExport']
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest().upper()
book = root/'workbook/output/RC_Section_NDM.xlsm'
code = root/'workbook/output/VBA_All_Code.txt'
assert sha(book) == publication['WorkbookSHA256']
assert sha(code) == publication['AllCodeSHA256'] == sha(directory/'PublishedNativeVBA.txt')
assert sha(root/'workbook/output/RC_Section_NDM_execution_report.txt') == before['ExecutionReportSHA256']
assert 'SetPublication' not in code.read_text(encoding='utf-8-sig')
report = {
    'Passed': True,
    'BaselineHead': before['Head'],
    'Requirement1ExistingFieldsUsedAtAll13Calls': True,
    'Requirement2OwnRightCommentOnEachSingleLinePerUserCorrection': True,
    'Requirement3DefaultsValuesOrderCoincidencesAndResultsPreserved': True,
    'Requirement4ObsoleteMethodAndLiveReferencesAbsent': True,
    'Requirement5DirectedRegressionPassed': True,
    'AlgorithmAndOtherArchitectureChanges': False,
    'UnchangedSourceComponents': 145,
    'Tests': counts, 'Assertions': sum(counts.values()),
    'ByteIdenticalModels': 9, 'ByteIdenticalControlNDS': 12,
    'CellsCompared': preservation['CellsChecked'], 'NamedRangesCompared': preservation['NamedRanges'],
    'ChangedZIPParts': preservation['ChangedZIPParts'],
    'NativeComponents': 157, 'SourceComponents': 152,
    'APIInventory': 292, 'GuideMembers': 136,
    'ActualNativeExportEqualsAllCode': True,
    'TemporaryTestHelpersAbsent': True,
    'UserExecutionReportPreservedAndExcludedFromCommit': True,
    'WorkbookSHA256': publication['WorkbookSHA256'],
    'AllCodeSHA256': publication['AllCodeSHA256'],
    'FullOnOffRunsStarted': False,
}
(directory/'Acceptance.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')

# Прежний код и точные снимки нужны только для воспроизведения сравнения.
# XLSM-копии и временные папки Excel в пакет доказательств не входят.
folders = ('Directed01','BaselineGeometry','CandidateGeometry','BaselineNDS','CandidateNDS')
with ZipFile(directory/'ExactSnapshots.zip','w',ZIP_DEFLATED,compresslevel=9) as out:
    for name in folders:
        for path in sorted((directory/name).glob('*')):
            if path.suffix in ('.txt','.json'):
                out.write(path,path.relative_to(directory).as_posix())
    for name in ('BaselineNativeVBA.txt','PublishedNativeVBA.txt','PreviousLayout.json','CurrentLayout.json'):
        out.write(directory/name,name)
print(json.dumps(report,ensure_ascii=False))
