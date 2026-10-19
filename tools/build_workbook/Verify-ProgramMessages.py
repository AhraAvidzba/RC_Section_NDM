"""Проверяет направленную приёмку сообщений и опубликованные артефакты."""
import hashlib
import json
from pathlib import Path
import re
import zipfile

ROOT = Path(__file__).resolve().parents[2]
EVIDENCE = ROOT / 'docs/regression/ProgramMessages_2026-10-19'


def digest(data):
    return hashlib.sha256(data).hexdigest().upper()


def receipt(name):
    return json.loads((EVIDENCE / name).read_text(encoding='utf-8-sig'))


def source_digest(path):
    # Git может менять CRLF/LF; содержимое кода, строк и комментариев сохраняется.
    return digest(path.read_text(encoding='utf-8-sig').encode('utf-8'))


def verify():
    acceptance = receipt('Acceptance.json')
    artifact = receipt('PublishedArtifact.json')
    equality = receipt('PublishedEquality.json')
    package = receipt('Package.json')
    review = receipt('Review.json')
    assert acceptance['Passed'] and artifact['Passed'] and equality['Passed']
    assert not acceptance['FullOnOffPerformed']
    assert acceptance['AcceptedSuites'] == 18 and acceptance['AssertionsFailed'] == 0
    assert len(review['Files']) == 112 and review['ProductionClasses'] == 97
    for item in review['Files']:
        path = ROOT / item['File']
        assert source_digest(path) == artifact['SourceTextSHA256'][path.stem], item['File']
    sources = {}
    for directory in ['src', 'tests']:
        for path in (ROOT / directory).rglob('*'):
            if path.suffix in ('.cls', '.bas'):
                assert path.stem not in sources, path
                sources[path.stem] = path
    assert equality['SourceComponents'] == len(sources) == 155
    assert equality['ExportComponents'] == 160
    assert not equality['Missing'] and not equality['Differences']
    assert not equality['UnexpectedCodeComponents']
    for item in equality['Components']:
        assert source_digest(sources[item['Name']]) == artifact['SourceTextSHA256'][item['Name']], item['Name']
    output = ROOT / 'workbook/output'
    assert digest((output / 'RC_Section_NDM.xlsm').read_bytes()) == artifact['WorkbookSHA256']
    code = (output / 'VBA_All_Code.txt').read_bytes()
    assert digest(code) == artifact['AllCodeSHA256'] == equality['ExportSHA256']
    assert len(re.findall(r'^COMPONENT: ', code.decode('utf-8-sig'), re.M)) == 160
    with zipfile.ZipFile(output / 'RC_Section_NDM.xlsm') as book:
        assert book.testzip() is None
        assert digest(book.read('xl/vbaProject.bin')) == artifact['VbaProjectSHA256']
        assert len(package['UnchangedParts']) == artifact['UnchangedPackageParts'] == 108
        for part in package['UnchangedParts']:
            assert digest(book.read(part['Part'])) == part['SHA256'], part['Part']
    total = 0
    with zipfile.ZipFile(EVIDENCE / 'Logs.zip') as archive:
        for test in acceptance['Tests']:
            raw = archive.read(test['Report'])
            assert digest(raw) == test['ReportSHA256']
            text = raw.decode('utf-8-sig')
            assert 'SOURCE_UNCHANGED: True' in text
            match = re.search(r'^===== ' + re.escape(test['Macro']) +
                              r' =====\r?\n(.*?)(?=^===== |\Z)', text, re.M | re.S)
            assert match, test
            body = match[1]
            assert re.search(r'^SUITE_FINISHED: ' + re.escape(test['Macro']) + ';', body, re.M)
            assert not re.search(r'^FAIL:|^RUNTIME ERROR|^SCRIPT ERROR|^WATCHDOG_FAILURE', body, re.M)
            totals = re.findall(r'^TOTAL[^\r\n]*passed=(\d+); failed=(\d+)', body, re.M)
            assert totals and all(int(failed) == 0 for _, failed in totals), test
            passed = sum(int(passed) for passed, _ in totals)
            assert passed == test['Passed'] and test['Failed'] == 0
            total += passed
    assert total == acceptance['AssertionsPassed'] == 80869
    assert receipt('CoreSourceContinuity.json')['Passed']
    assert receipt('ClassComments.json')['status'] == 'PASS'
    return {'Passed': True, 'Scope': 'directed messages acceptance',
            'ProductionComponents': 112, 'Suites': 18, 'Assertions': total,
            'FullOnOffPerformed': False, 'WorkbookSHA256': artifact['WorkbookSHA256']}


if __name__ == '__main__':
    print(json.dumps(verify(), ensure_ascii=False))
