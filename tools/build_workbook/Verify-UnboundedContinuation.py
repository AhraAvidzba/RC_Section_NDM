"""Проверяет линейное продолжение, связанные тесты и сохранность готовой книги."""
import argparse
import hashlib
import importlib.util
import json
import re
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
D = ROOT / 'docs/regression/ExcelGuardGuide_2026-10-13/ExtensionLimit100_2026-10-16'
spec = importlib.util.spec_from_file_location('continuation_reader', Path(__file__).with_name('Verify-SectionRecipeGuide.py'))
reader = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reader)


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def receipt(name):
    return json.loads(payload(D / name).decode('utf-8-sig'))


def payload(path):
    if path.exists():
        return path.read_bytes()
    manifest = json.loads((D / 'Archive.json').read_text(encoding='utf-8'))
    archive = D / 'Evidence.zip'
    assert sha(archive) == manifest['SHA256']
    name = path.relative_to(D).as_posix()
    with ZipFile(archive) as data:
        result = data.read(name)
    assert hashlib.sha256(result).hexdigest().upper() == manifest['Entries'][name]
    return result


def evidence_text(path):
    return payload(path).decode('utf-8-sig')


def components(path):
    text = evidence_text(path).replace('\r\n', '\n')
    pattern = r'(?ms)^={100}\nCOMPONENT: ([^\n]+)\nTYPE: (\d+)\nLINES: (\d+)\n={100}\n(.*?)(?=^={100}\nCOMPONENT: |\Z)'
    return {m[1]: (m[2], m[4].strip()) for m in re.finditer(pattern, text)}


def normalized(body):
    parts = re.findall(r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27\n]+|\n', body)
    return ''.join(part if part.startswith(('"', "'")) else part.lower() for part in parts).strip()


def scope():
    before = components(D / 'BaselineNativeVBA.txt')
    baseline_code = components(D / 'BaselineAllCode.txt')
    assert before.keys() == baseline_code.keys()
    assert all(normalized(before[name][1]) == normalized(baseline_code[name][1]) for name in before), 'Baseline VBA differs from saved ALL CODE'
    after = components(D / 'FinalCandidateNativeVBA.txt')
    assert before.keys() == after.keys() and len(after) == 158
    changed = sorted(name for name in before if normalized(before[name][1]) != normalized(after[name][1]))
    assert changed == sorted(('CMaterialDiagram', 'CMaterialModelProvider', 'CNDMResultsWriter',
                              'modTestMaterialDiagrams', 'modTestBatchCalculation', 'modTestSectionSolver')), changed
    assert receipt('FinalCandidateSourceEquality.json')['Passed']
    return changed


def fixture_equivalence():
    old = components(D / 'CandidateNativeVBA.txt')
    new = components(D / 'FinalCandidateNativeVBA.txt')
    assert old.keys() == new.keys()
    for name in old:
        if name != 'modTestBatchCalculation':
            assert normalized(old[name][1]) == normalized(new[name][1]), name
    pattern = r'(?ms)^(?:Public Function RunBatchCalculationTests\(.*?^End Function|Private Sub RestoreBatchSuiteProfileDefaults\(.*?^End Sub)'
    assert len(re.findall(pattern, old['modTestBatchCalculation'][1])) == 2
    assert len(re.findall(pattern, new['modTestBatchCalculation'][1])) == 2
    assert normalized(re.sub(pattern, '', old['modTestBatchCalculation'][1])) == normalized(re.sub(pattern, '', new['modTestBatchCalculation'][1]))
    current_entry = re.findall(pattern, new['modTestBatchCalculation'][1])[0]
    assert 'SetProfileValue "Calculation.Crack.Width", "PR1", "No"' in current_entry
    current_restore = re.findall(pattern, new['modTestBatchCalculation'][1])[1]
    assert 'SetProfileValue "Calculation.Crack.Width", "PR1", originalPr1CrackWidth' in current_restore
    text = evidence_text(D / 'BaselineProfileContext/modTestBatchCalculation.ProbeBatchProfileAssumptions.txt')
    assert 'phase=Saved; passed=105; failures=9' in text
    assert 'phase=Normalized; passed=114; failures=0' in text
    return dict(Passed=True, RuntimeComponentsUnchanged=True,
                OnlyChangedTestProcedures=['RunBatchCalculationTests', 'TestRestoreBatchSuiteProfileDefaults'])


def help_replacements(book):
    catalog = (ROOT / 'tools/build_workbook/SettingsCatalog.ps1').read_text(encoding='utf-8-sig')
    texts = [m[1] for m in re.finditer(r'"([^"\r\n]+)"', catalog)]
    pairs = [('General.DiagramExtension = Yes', 'General.DiagramExtension = Yes линейно'),
             ('Yes разрешает численно продолжать', 'Yes линейно продолжает активные ветви'),
             ('Настройка применяется к равновесию прочности', 'Настройка действует при расчёте прочности')]
    result = {}
    for old_prefix, new_prefix in pairs:
        found = [(address, cell['value']) for address, cell in book['sheets']['Справка']['cells'].items()
                 if (cell['value'] or '').startswith(old_prefix)]
        new = [text for text in texts if text.startswith(new_prefix)]
        assert len(found) == len(new) == 1, (old_prefix, found, new)
        result[found[0][0]] = (found[0][1], new[0])
    return result


def package():
    changed = scope()
    old = reader.read(D / 'Baseline.xlsm')
    replacements = help_replacements(old)
    equivalence = fixture_equivalence()
    accepted = D / 'AcceptedFinal.xlsm'
    assert not accepted.exists(), 'Accepted package already exists'
    with ZipFile(D / 'Baseline.xlsm') as source, ZipFile(D / 'FinalCandidate/Continuation.xlsm') as candidate, ZipFile(accepted, 'w') as output:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == 'xl/vbaProject.bin':
                data = candidate.read(info.filename)
            elif info.filename == 'xl/sharedStrings.xml':
                for old_text, new_text in replacements.values():
                    needle = escape(old_text).encode('utf-8')
                    assert data.count(needle) == 1
                    data = data.replace(needle, escape(new_text).encode('utf-8'))
            output.writestr(info, data)
    with ZipFile(D / 'Baseline.xlsm') as source, ZipFile(accepted) as output:
        assert source.namelist() == output.namelist()
        parts = sorted(n for n in source.namelist() if source.read(n) != output.read(n))
        assert parts == ['xl/sharedStrings.xml', 'xl/vbaProject.bin'], parts
    result = dict(Passed=True, ChangedComponents=changed, ChangedZIPParts=parts,
                  HelpChanges=replacements, AcceptedSHA256=sha(accepted),
                  TestedWorkbookSHA256=sha(D / 'FinalCandidate/Continuation.xlsm'),
                  EquivalentEarlierTestSourceSHA256=sha(D / 'Unbounded/Continuation.xlsm'), FixtureEquivalence=equivalence)
    with ZipFile(accepted) as output, ZipFile(D / 'FinalCandidate/Continuation.xlsm') as candidate:
        assert output.read('xl/vbaProject.bin') == candidate.read('xl/vbaProject.bin')
        result['CompiledVbaSHA256'] = hashlib.sha256(output.read('xl/vbaProject.bin')).hexdigest().upper()
    (D / 'FinalPackage.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k: v for k, v in result.items() if k != 'HelpChanges'}))


def native_run(stage, macros, counts=None):
    run = receipt(stage + '/Run.json')
    assert run['Passed'] and run['ProductionUnchanged'], stage
    assert [r['Macro'] for r in run['Runs']] == macros, stage
    package_data = receipt('FinalPackage.json')
    if run['SourceSHA256'] == package_data['EquivalentEarlierTestSourceSHA256']:
        assert package_data['FixtureEquivalence']['Passed']
        assert 'modTestBatchCalculation.RunBatchCalculationTests' not in macros
    else:
        assert run['SourceSHA256'] == package_data['TestedWorkbookSHA256'], stage
    assertions = 0
    texts = []
    for macro in macros:
        text = evidence_text(D / stage / (macro + '.txt'))
        assert not re.search(r'(?m)^FAIL:|RUNTIME ERROR:|^ERROR:', text), stage
        totals = re.findall(r'(?m)^TOTAL[^\r\n]*passed=(\d+); failed=(\d+)', text)
        assert totals and all(int(p) > 0 and int(f) == 0 for p, f in totals), (stage, totals)
        assertions += int(totals[-1][0])
        texts.append(text)
    if counts:
        for text, expected in zip(texts, counts):
            assert expected in text, (stage, expected)
    return dict(Stage=stage, Assertions=assertions, Seconds=sum(r['Seconds'] for r in run['Runs'])), texts


def tests():
    runs = []
    focused = ['modTestMaterialDiagrams.RunMaterialDiagramTests', 'modTestBatchCalculation.RunPlateauResultsDiagramTests',
               'modTestBatchCalculation.RunUnboundedDiagramResultsTests', 'modTestSectionSolver.RunUnboundedExtensionArithmeticTests',
               'modTestSectionSolver.RunSectionSolverTests']
    runs.append(native_run('UnboundedFocused', focused)[0])
    # The batch suite completed successfully; the following capacity attempt used
    # the user's PointedSection instead of that test's explicit Rectangles fixture.
    # Accept only the completed batch suite from that stopped combined invocation.
    combined = receipt('CoreRegressionFinal/Run.json')
    assert combined['ProductionUnchanged'] and not combined['Passed']
    assert combined['SourceSHA256'] == receipt('FinalPackage.json')['TestedWorkbookSHA256']
    assert [r['Macro'] for r in combined['Runs']] == ['modTestBatchCalculation.RunBatchCalculationTests', 'modTestCapacitySolver.RunCapacitySolverTests']
    assert combined['Runs'][0]['Passed'] and not combined['Runs'][1]['Passed']
    batch_text = evidence_text(D / 'CoreRegressionFinal/modTestBatchCalculation.RunBatchCalculationTests.txt')
    assert not re.search(r'(?m)^FAIL:|RUNTIME ERROR:', batch_text)
    assert 'TOTAL_BATCH: passed=30982; failed=0;' in batch_text
    capacity_attempt = evidence_text(D / 'CoreRegressionFinal/modTestCapacitySolver.RunCapacitySolverTests.txt')
    assert re.findall(r'(?m)^FAIL:[^\r\n]+', capacity_attempt) == ['FAIL: capacity.tfDefaults.geometry; actual=PointedSection; expected=Rectangles']
    runs.append(dict(Stage='CoreRegressionFinal/BatchOnly', Assertions=30982, Seconds=combined['Runs'][0]['Seconds'],
                     CombinedInvocationPassed=False, ExcludedFollowingCapacityAttempt=True))
    capacity = evidence_text(D / 'CapacityFixture.txt')
    assert not re.search(r'(?m)^FAIL:|RUNTIME ERROR:|WATCHDOG_FAILURE:', capacity)
    assert 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' in capacity
    assert 'EXPLICIT_GEOMETRY_FIXTURE: original=PointedSection; Geometry.Type=Rectangles; fixture only, source unchanged' in capacity
    assert 'SOURCE_SHA256: ' + receipt('FinalPackage.json')['TestedWorkbookSHA256'] in capacity
    totals = re.findall(r'TOTAL_CAPACITY: passed=(\d+); failed=0; elapsedSec=([\d.]+)', capacity)
    assert len(totals) == 1
    runs.append(dict(Stage='CapacityFixture', Assertions=int(totals[0][0]), Seconds=float(totals[0][1])))
    runs.append(native_run('RemainingCore', ['modTestCrackWidth.RunCrackWidthTests', 'modTestBatchCalculation.RunAudit02OnOffComparisonTests'])[0])
    for shape in ('CircleSym', 'RectL', 'RoundedTapered', 'HollowOffset', 'ImportedFixture'):
        macros = ['modTestBatchCalculation.RunAudit03ExtendedOverloadEquilibriumTests',
                  'modTestBatchCalculation.RunAudit03ExtendedStressFeasibilityTests']
        run, texts = native_run('Geometry_' + shape, macros, ['knownStates=96;', 'variant=Newton; cases=48;'])
        assert len(re.findall(r'(?m)^EXTENDED_STRESS:', texts[1])) == 48
        assert not re.search(r'(?m)^EXTENDED_STRESS:[^\r\n]*\|converged=False', texts[1])
        runs.append(run)
        run, texts = native_run('Secant_' + shape, macros[1:], ['variant=Secant; cases=48;'])
        assert len(re.findall(r'(?m)^EXTENDED_STRESS:', texts[0])) == 48
        assert not re.search(r'(?m)^EXTENDED_STRESS:[^\r\n]*\|converged=False', texts[0])
        runs.append(run)
    saved = {}
    for stage in ('SavedOn', 'SavedOff', 'QuietTimingOn', 'QuietTimingOff'):
        run = receipt(stage + '/Run.json')
        assert run['Passed'] and run['ProductionUnchanged']
        assert len(run['Runs']) == 1
        text = evidence_text(D / stage / 'modTestBatchCalculation.RunSavedStabilityDiagnostic.txt')
        rows = [line.split('\t') for line in text.splitlines()]
        states = {r[1]: r for r in rows if r[0] == 'STATE'}
        for name in ('StrengthState', 'CrackedState'):
            r = states[name]
            assert r[2:6] == ['True', 'FAIL', 'False', 'True'], (stage, name, r)
            assert abs(float(r[12])) < 1 and abs(float(r[13])) < 1000 and abs(float(r[14])) < 1000
        case = next(r for r in rows if r[0] == 'CASE')
        saved[stage] = dict(Seconds=float(case[2]), Iterations={n: int(float(states[n][19])) for n in ('StrengthState', 'CrackedState')},
                            SteelMaximum={n: float(states[n][18]) for n in ('StrengthState', 'CrackedState')})
    assert min(saved['SavedOn']['SteelMaximum'].values()) > 10
    assert saved['QuietTimingOn']['Iterations'] == saved['SavedOn']['Iterations']
    assert saved['QuietTimingOff']['Iterations'] == saved['SavedOff']['Iterations']
    assert saved['QuietTimingOn']['Seconds'] < 10 and saved['QuietTimingOff']['Seconds'] < 10
    return runs, saved


def preservation():
    package_data = receipt('FinalPackage.json')
    before, after = reader.read(D / 'Baseline.xlsm'), reader.read(D / 'AcceptedFinal.xlsm')
    assert before['names'] == after['names'] and before['drawings'] == after['drawings']
    assert before['sheets'].keys() == after['sheets'].keys()
    checked = 0
    for name, sheet in before['sheets'].items():
        actual = after['sheets'][name]
        assert sheet['rows'] == actual['rows'], name
        assert reader.canonical(sheet['document']) == reader.canonical(actual['document']), name
        assert sheet['cells'].keys() == actual['cells'].keys()
        for address, old in sheet['cells'].items():
            new = actual['cells'][address]
            checked += 1
            if name == 'Справка' and address in package_data['HelpChanges']:
                assert [old['value'], new['value']] == package_data['HelpChanges'][address]
                assert all(old[k] == new[k] for k in ('formula', 'style', 'attrs'))
            else:
                assert old == new, (name, address)
    return dict(Cells=checked, NamedRanges=len(before['names']), HelpCells=sorted(package_data['HelpChanges']))


def accept(published):
    changed = scope()
    equivalence = fixture_equivalence()
    runs, saved = tests()
    preserved = preservation()
    assert receipt('FinalCandidateSourceEquality.json')['Passed']
    marker = re.compile(r'NDM\.RECIPE\.GUIDE/\d|\b[0-9a-f]{64}\b', re.I)
    assert not marker.search(evidence_text(D / 'FinalCandidateNativeVBA.txt'))
    with ZipFile(D / 'AcceptedFinal.xlsm') as archive:
        assert hashlib.sha256(archive.read('xl/vbaProject.bin')).hexdigest().upper() == receipt('FinalPackage.json')['CompiledVbaSHA256']
        assert all(not marker.search(archive.read(n).decode('utf-8-sig')) for n in archive.namelist() if n.endswith(('.xml', '.rels')))
    if published:
        assert sha(ROOT / 'workbook/output/RC_Section_NDM.xlsm') == receipt('FinalPackage.json')['AcceptedSHA256']
        assert receipt('PublishedSourceEquality.json')['Passed']
        assert components(D / 'PublishedNativeVBA.txt') == components(D / 'FinalCandidateNativeVBA.txt')
        assert (ROOT / 'workbook/output/VBA_All_Code.txt').read_bytes() == payload(D / 'PublishedNativeVBA.txt')
        assert sha(ROOT / 'workbook/output/RC_Section_NDM_execution_report.txt') == receipt('Before.json')['Hashes']['workbook/output/RC_Section_NDM_execution_report.txt']
    result = dict(Passed=True, PublishedVerified=published, ChangedComponents=changed,
                  Runs=runs, TotalAssertions=sum(r['Assertions'] for r in runs), SavedExample=saved, Preserved=preserved,
                  FixtureEquivalence=equivalence)
    (D / 'Acceptance.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(result, ensure_ascii=False))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--package', action='store_true')
    parser.add_argument('--published', action='store_true')
    args = parser.parse_args()
    package() if args.package else accept(args.published)
