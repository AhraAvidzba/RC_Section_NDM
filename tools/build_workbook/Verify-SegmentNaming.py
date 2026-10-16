"""Адресная приёмка явных имён: исходные данные, геометрия, происхождение и реальные схемы."""
import argparse
import hashlib
import json
import re
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
D = ROOT / 'docs/regression/ExcelGuardGuide_2026-10-13/SegmentNaming_2026-10-16'
MATRICES = ['DumpCircleRecipeMatrix', 'DumpRectSetRecipeMatrix', 'DumpRoundedRecipeMatrix',
            'DumpHollowRecipeMatrix', 'DumpHollowRecipeEdgeMatrix', 'DumpRecipeRuleContracts']
PLOTS = ['DumpCircleRecipePreview', 'DumpRectSetRecipePreview', 'DumpRoundedRecipePreview', 'DumpHollowRecipePreview']


def data(name):
    path = D / name
    if path.exists():
        return path.read_bytes()
    with ZipFile(D / 'Evidence.zip') as archive:
        return archive.read(name)


def text(name):
    # Windows PowerShell 5 и PowerShell 7 различаются BOM; содержимое и биты Double не нормализуем.
    return data(name).decode('utf-8-sig').replace('\r\n', '\n')


def receipt(name):
    return json.loads(text(name))


def compare(before, after):
    a, b = text(before), text(after)
    assert a == b, f'Snapshot differs: {before} / {after}'
    assert not re.search(r'(?m)^FAIL|^ERROR:|RUNTIME ERROR', a), before
    return {'Before': before, 'After': after, 'Characters': len(a), 'Rows': len(a.splitlines()),
            'ContentSHA256': hashlib.sha256(a.encode('utf-8')).hexdigest()}


def verify(published=False):
    comparisons = [compare(f'BaselineMatrices/{name}.txt', f'CandidateMatrices/{name}.txt') for name in MATRICES]
    comparisons.append(compare('BaselinePaths/DumpRecipePathContracts.txt', 'CandidatePaths/DumpRecipePathContracts.txt'))
    paths = text('CandidatePaths/DumpRecipePathContracts.txt')
    cases = re.findall(r'(?m)^CASE\|', paths)
    assert len(cases) == 60
    for stage in ['BaselineMatrices', 'CandidateMatrices', 'BaselinePaths', 'CandidatePaths', 'BaselinePlots', 'CandidatePlots', 'DirectedAccepted']:
        run = receipt(f'{stage}/Run.json')
        assert run['Passed'] and run['ProductionUnchanged'], stage
        assert all(item['Passed'] for item in run['Runs']), stage
    checks = 0
    suites = receipt('DirectedAccepted/Run.json')['Runs']
    assert len(suites) == 12
    for suite in suites:
        body = text('DirectedAccepted/'+suite['Macro']+'.txt')
        total = re.findall(r'(?m)^TOTAL[^\n]*passed=(\d+); failed=(\d+)\b', body)
        assert len(total) == 1 and int(total[0][1]) == 0 and not re.search(r'(?m)^FAIL|RUNTIME ERROR', body)
        checks += int(total[0][0])
    for name in PLOTS:
        comparisons.append(compare(f'BaselinePlots/{name}.txt', f'CandidatePlots/{name}.txt'))
    pngs = ['Circle', 'Rectangle', 'LSection', 'TwoRectangles', 'RoundedSimple', 'RoundedTapered',
            'RoundedMixed', 'RoundedMixedLeft', 'HollowCentered', 'HollowOffset']
    for name in pngs:
        assert data(f'BaselinePlots/{name}.png') == data(f'CandidatePlots/{name}.png'), name
    assert receipt('CandidateSourceEquality.json')['Passed']
    assert receipt('SourceScope.json')['Passed'] and receipt('SourceScope.json')['NewComponents'] == 0
    equality = receipt('PublishedSourceEquality.json' if published else 'CandidateSourceEquality.json')
    sources = {path.stem:path for folder in ['src','tests'] for path in (ROOT/folder).rglob('*') if path.suffix in ('.bas','.cls')}
    assert len(sources) == equality['SourceComponents'] == 153 and equality['ExportComponents'] == 158
    for item in equality['Components']:
        assert item['Equal'] and hashlib.sha256(sources[item['Name']].read_bytes()).hexdigest().upper() == item['SourceSHA256']
    assert receipt('ClassComments.json')['status'] == 'PASS'
    guide = (ROOT/'src/Common/modSectionRecipeApiGuide.bas').read_text(encoding='utf-8-sig')
    assert all(not line.strip() or line.startswith(('Attribute VB_Name = ', "'")) for line in guide.splitlines())
    before = receipt('Before.json')
    report = ROOT/'workbook/output/RC_Section_NDM_execution_report.txt'
    assert hashlib.sha256(report.read_bytes()).hexdigest().upper() == before['Hashes']['workbook/output/RC_Section_NDM_execution_report.txt']
    if published:
        publication = receipt('Publication.json')
        assert publication['Status'] == 'PUBLISHED_VERIFIED'
        assert publication['UserDataBaselineSHA256'] == receipt('Package.json')['BaselineSHA256']
        assert receipt('Package.json')['UserLiveVBAUnchanged']
        for name in ['PublishedSourceEquality.json', 'PublishedGuidePresentation.json']:
            assert receipt(name)['Passed'], name
        assert (ROOT/'workbook/output/VBA_All_Code.txt').read_bytes() == data('PublishedNativeVBA.txt')
        assert hashlib.sha256((ROOT/'workbook/output/RC_Section_NDM.xlsm').read_bytes()).hexdigest() == publication['WorkbookSHA256']
    result = {'Status': 'SEGMENT_NAMING_PASS', 'Suites': len(suites), 'Passed': checks, 'Failed': 0,
              'PathCases': len(cases), 'Paths': len(re.findall(r'(?m)^PATH\|', paths)),
              'Segments': len(re.findall(r'(?m)^SEGMENT\|', paths)), 'IdenticalPNGs': len(pngs),
              'Comparisons': comparisons, 'FullOnOffRun': False, 'PublishedVerified': published}
    (D/'Acceptance.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps({k:v for k,v in result.items() if k != 'Comparisons'}, ensure_ascii=False))


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--published', action='store_true')
    args = parser.parse_args()
    verify(args.published)
