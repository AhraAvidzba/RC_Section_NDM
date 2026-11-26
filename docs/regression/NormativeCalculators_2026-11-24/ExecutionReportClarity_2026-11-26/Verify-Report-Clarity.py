"""ДЛЯ ТЕСТОВ: принимает готовые отчеты и неизменность физических результатов по требованиям задачи."""
import hashlib
import json
from pathlib import Path
import re

HERE = Path(__file__).resolve().parent
ROOT = next(p for p in HERE.parents if (p / 'AGENTS.md').is_file())
def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))
def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()
def report(label, case):
    path = HERE / f'{label}_{case}.txt'
    raw = path.read_bytes()
    return raw.decode('utf-16') if raw.startswith((b'\xff\xfe', b'\xfe\xff')) else raw.decode('utf-8-sig')

proof = read(HERE / 'PackageProof.json')
assert proof['Status'] == 'PASS' and proof['NativeComponents'] == 209 and proof['SourceEquality']
assert sha(HERE / 'Candidate.xlsm') == proof['CandidateSHA256']
macros = ['modTestWorkbookInterface.RunExecutionReportFormattingTests',
          'modTestBatchCalculation.RunAudit03FormationReportTests',
          'modTestBatchCalculation.RunZeroLoadCombinationTests',
          'modTestBatchCalculation.RunNormativeStabilityTests',
          'modTestTorsionPost.RunTorsionOnlyTests']
runs = read(HERE / 'RelatedRuns_Publication.json')
assert len(runs) == 2 and [r['Mode'] for r in runs] == ['No', 'Yes']
assertions = 0
for run in runs:
    path = ROOT / run['Report']
    body = path.read_text(encoding='utf-8-sig')
    assert run['Passed'] and sha(path) == run['SHA256'] and run['Macros'] == macros
    assert re.findall(r'^SUITE_STARTED: ([^;]+);', body, re.M) == macros
    assert re.findall(r'^SUITE_FINISHED: ([^;]+);', body, re.M) == macros
    assert 'SOURCE_SHA256: ' + proof['CandidateSHA256'] in body
    assert 'SOURCE_UNCHANGED: True' in body and 'WATCHDOG_COMPLETED: exit=0;' in body
    assert not re.search(r'^FAIL\b|[Ff]ailed=[1-9]|^RUNTIME[: ]|WATCHDOG_FAILURE', body, re.M)
    totals = re.findall(r'^.*?: [Pp]assed=(\d+); [Ff]ailed=(\d+)', body, re.M)
    assert len(totals) == 5 and all(int(f) == 0 for _, f in totals)
    assertions += sum(int(p) for p, _ in totals)
    assert 'reader.Traw.emptyProfile' in body and 'reader.Tfiltered.emptyProfile' in body
    assert read(Path(str(path) + '.ProcessClosure.json'))['Passed']

cases = ['UserExample', 'TOnlyBlankProfile', 'SP63Stability', 'SP35Width', 'SP63Width', 'CompressedNMT']
observations = read(HERE / 'SavedExamples.json')
assert len(observations) == 12
index = {(r['Label'], r['Case']): r for r in observations}
assert len(index) == 12 and set(index) == {(label, case) for label in ['Baseline', 'Candidate'] for case in cases}
physics = []
for case in cases:
    before, after = index['Baseline', case], index['Candidate', case]
    assert before['SourceSHA256'] == proof['BaselineSHA256']
    assert after['SourceSHA256'] == proof['CandidateSHA256']
    for key in ['Elements', 'Summaries', 'Counters']:
        assert before[key] == after[key], (case, key)
    physics.append({'Case': case, 'ElementRows': len(after['Elements']) - 1,
                    'ElementsExactlyEqual': True, 'AllFourSummaryTablesExactlyEqual': True,
                    'SolverCountersExactlyEqual': True})

text = report('Candidate', 'UserExample')
sections = ['Поиск несущей способности', 'Текущее НДС по модели прочности',
            'Образование нормальной трещины', 'Текущее НДС для проверки трещин',
            'Проверка продольных трещин', 'Ширина раскрытия нормальных трещин',
            'Статистика всего пакетного расчета', 'Кручение и проверочная сумма напряжений NM+T', 'Запись результатов']
positions = [text.index('\n' + title + '\n') if '\n' + title + '\n' in text else text.index('\r\n' + title + '\r\n') for title in sections]
assert positions == sorted(positions)
for sentence in ['внутренняя расчетная система координат элементов сечения',
                 'N=-50 tf; Mx=40 tf*m; My=30 tf*m',
                 'N=-50 tf; Mx=14.605263157895 tf*m; My=12.105263157895 tf*m',
                 'X=357.894736842105 mm; Y=507.894736842105 mm',
                 'N, Nint, RN — Н', 'RMx, RMy — Н*мм', 'O(0; 0)',
                 'Статус текущего НДС по прочности (StrengthState)',
                 'Статус текущего НДС для трещин (CrackedState)',
                 'по действующему фильтру', 'проба №=',
                 'Длительность ведения журнала', 'включая подготовку до старта журнала']:
    assert sentence in text, sentence
assert 'Mx расчетный:' not in text and 'Статус прямого НДС:' not in text
assert 'траектория является чисто осевой' not in text
assert 'точка lambda=' not in text
assert not re.search(r'Плоскость [12]: Ncr: 0(?:\s|$)|D=0(?:\s|$)', text)
assert len(re.findall(r'Плоскость [12]: Nult,stab:', text)) == 2
assert 'Ncr и жесткость D не рассчитываются' in text
assert not re.search(r'\] (?:As|Abt|ds_eq|ls|psi_s|a_crc):', text)

only_t = report('Candidate', 'TOnlyBlankProfile')
for sentence in ['Полностью исключено сочетаний с малыми усилиями и без кручения: 0',
                 'Принято сочетаний для обработки: 3',
                 'Из них запрошено только T (кручение), без НДМ: 1',
                 'Профиль НДМ: не требуется: сочетание обрабатывается только по T',
                 'Сочетание не исключено.', 'Общий статус проверки T/NMT: OK']:
    assert sentence in only_t, sentence
assert index['Candidate', 'TOnlyBlankProfile']['Counters'] == index['Candidate', 'UserExample']['Counters']
assert 'не ограничен по растяжению / OK' in report('Candidate', 'CompressedNMT')
sp63_stability = report('Candidate', 'SP63Stability')
assert 'Плоскость 1: Ncr:' in sp63_stability and 'Плоскость 1: phi_l и D:' in sp63_stability
assert 'ядровое расстояние r' not in sp63_stability and 'Nult,stab:' not in sp63_stability
sp35, sp63 = report('Candidate', 'SP35Width'), report('Candidate', 'SP63Width')
assert 'Rr:' in sp35 and 'psi по СП 35:' in sp35 and '] psi_s:' not in sp35
assert '] psi_s:' in sp63 and '] Abt:' in sp63 and '] ls:' in sp63
scope = read(HERE / 'MethodChoiceScope.json')
assert scope['NumericalAlgorithmsUnchanged'] and all(scope['OnlyStringLiteralsChanged'].values())
result = {'Status': 'PASS', 'RelatedExecutions': 10, 'RelatedAssertions': assertions,
          'SavedExampleExecutions': 12, 'PhysicalComparisons': physics,
          'SeparateSections': sections, 'EffortReferencePointsAndCoordinateSystem': True,
          'InternalDiagnosticUnitsAndProbeNumbers': True, 'InapplicableParametersNotInvented': True,
          'ProfileOptionalForTOnly': True, 'TOnlyNotCountedAsFiltered': True,
          'FilteredCountExcludesEmptyRows': True, 'NoAdditionalNDSSolvesForTOnly': True,
          'FullRunStarted': False, 'CandidateSHA256': proof['CandidateSHA256'],
          'Review01Note': 'Case-sensitive SP35 branch rendering detected during report review; fixed, final reports explicitly checked.',
          'ThresholdsTolerancesAndPhysicsUnchanged': True}
(HERE / 'Acceptance.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps(result, ensure_ascii=False))
