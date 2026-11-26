"""ДЛЯ ТЕСТОВ: сверяет установленную книгу и конечный отчет с принятой частной пробой."""
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

HERE = Path(__file__).resolve().parent
ROOT = next(p for p in HERE.parents if (p / 'AGENTS.md').is_file())
def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))
def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()

proof = read(HERE / 'PackageProof.json')
assert read(HERE / 'Acceptance.json')['Status'] == 'PASS'
book = ROOT / 'workbook/output/RC_Section_NDM.xlsm'
export = ROOT / 'workbook/output/VBA_All_Code.txt'
report = ROOT / 'workbook/output/RC_Section_NDM_execution_report.txt'
assert sha(book) == proof['CandidateSHA256']
assert sha(export) == proof['ExportSHA256']
published = read(HERE / 'PublishedRun.json')
if isinstance(published, dict):
    published = [published]  # Windows PowerShell сериализует единственную запись как объект.
assert len(published) == 1
published = published[0]
candidate = next(r for r in read(HERE / 'SavedExamples.json') if r['Label'] == 'Candidate' and r['Case'] == 'UserExample')
for key in ['Elements', 'Summaries', 'Counters']:
    assert published[key] == candidate[key], key
assert published['SourceSHA256'] == proof['CandidateSHA256']
assert report.read_bytes() == (HERE / 'Published_UserExample.txt').read_bytes()
text = report.read_text(encoding='utf-16')
assert 'Книга: RC_Section_NDM.xlsm' in text
assert 'Папка книги: ' + str(book.parent) in text
assert 'Количество: 1.' in text and 'ID: 1.' in text
assert 'N=-50 tf; Mx=40 tf*m; My=30 tf*m' in text
for section in read(HERE / 'Acceptance.json')['SeparateSections']:
    assert '\n' + section + '\n' in text
with ZipFile(book) as installed, ZipFile(HERE / 'Baseline.xlsm') as original:
    assert installed.namelist() == original.namelist()
    assert [n for n in installed.namelist() if installed.read(n) != original.read(n)] == ['xl/vbaProject.bin']
delivery = {'Status': 'PASS', 'WorkbookSHA256': sha(book), 'ExportSHA256': sha(export),
            'ReportSHA256': sha(report), 'ReportGeneratedFromInstalledWorkbook': True,
            'UserDataAndFormattingPreserved': True, 'PublishedPhysicalResultsEqualAcceptedCandidate': True,
            'WorkbookUnchangedByVerification': True, 'FullRunStarted': False}
(HERE / 'Delivery.json').write_text(json.dumps(delivery, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
files = sorted(p for p in HERE.iterdir() if p.is_file() and p.suffix not in ['.zip'] and p.name != 'EvidenceInventory.json')
inventory = [{'Name': p.name, 'Bytes': p.stat().st_size, 'SHA256': sha(p),
              'Archived': p.suffix != '.xlsm' and not p.name.startswith(('NativeCandidate', 'BaselineNative')) and p.name != 'VBA_All_Code.txt'} for p in files]
(HERE / 'EvidenceInventory.json').write_text(json.dumps(inventory, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
with ZipFile(HERE / 'Evidence.zip', 'w', ZIP_DEFLATED) as archive:
    for record in inventory:
        if record['Archived']:
            archive.write(HERE / record['Name'], record['Name'])
    archive.write(HERE / 'EvidenceInventory.json', 'EvidenceInventory.json')
print(json.dumps(delivery, ensure_ascii=False))
