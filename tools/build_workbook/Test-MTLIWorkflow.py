"""ДЛЯ ТЕСТОВ: точность итоговых таблиц и время полной команды MTLI.

Сохраняет все превышения 2% и изменения текстовых полей. Различный метод
или путь не приравнивается к отказу. Первичные файлы не изменяются.
"""
import argparse
import json
import re
import statistics
from pathlib import Path

TABLES = {'Strength.Data': 'Strength.Headers', 'Cracks.Data': 'Cracks.Headers', 'Stability.Data': 'Stability.Labels'}
NUMERIC_TYPES = {'2', '3', '4', '5', '6', '14', '17'}


def read_cells(path):
    """Читает экспорт с исходными типами и адресами каждой ячейки."""
    cells = {}
    for line in path.read_text(encoding='utf-8-sig').splitlines():
        if line.startswith('CELL|'):
            _, table, row, col, kind, value = line.split('|', 5)
            cells[table, int(row), int(col)] = kind, value
    return cells


def floor_for(header):
    """Абсолютный порог для практически нулевых величин, в единицах вывода.

    Пороги фиксированы до сравнения и существенно меньше точности вывода.
    Не предназначены для разрешения процентов ошибки у значимых величин.
    """
    if 'κ' in header:
        return 1e-12
    if 'ε' in header:
        return 1e-10
    if 'MPa' in header:
        return 1e-5
    if 'tf*m' in header or 'tf' in header:
        return 1e-6
    if 'mm' in header:
        return 1e-6
    if 'm2' in header:
        return 1e-10
    return 1e-8


def compare(reference, candidate):
    """Все числовые результаты сравниваются; диагностические строки отделены."""
    base, current = read_cells(reference), read_cells(candidate)
    differences, strings, maxima, checked = [], [], {}, 0
    for key, (kind, value) in base.items():
        table, row, column = key
        if table not in TABLES:
            continue
        other_kind, other = current.get(key, ('MISSING', ''))
        header_row = 1 if table == 'Stability.Data' else 4
        header = base.get((TABLES[table], header_row, column), ('', ''))[1]
        if kind in NUMERIC_TYPES and other_kind in NUMERIC_TYPES:
            actual, expected = float(other.replace(',', '.')), float(value.replace(',', '.'))
            delta = abs(actual - expected)
            absolute = floor_for(header)
            percentage = 100 * delta / abs(expected) if abs(expected) > absolute else None
            checked += 1
            if percentage is not None:
                maxima[table] = max(maxima.get(table, 0), percentage)
            if delta > absolute + 0.02 * abs(expected):
                differences.append(dict(Table=table, Row=row, Column=column, Header=header, Reference=expected, Actual=actual, ErrorPct=percentage, AbsoluteFloor=absolute))
        elif (kind, value) != (other_kind, other):
            strings.append(dict(Table=table, Row=row, Column=column, Header=header, Reference=value, Actual=other))
    return dict(WithinTwoPercent=not differences, NumericCells=checked, MaxErrorPct=maxima, Exceedances=differences, ChangedText=strings)


def analyse(baseline, candidate):
    """Сопоставляет каждый повтор с исходным релизным VBA того же случая."""
    comparisons, timings = [], {}
    for path in sorted(candidate.glob('Workflow_*_*_*.txt')):
        match = re.fullmatch(r'Workflow_(\d+)_(\d+)_(\d+)\.txt', path.name)
        case, variant, repeat = map(int, match.groups())
        refs = sorted(baseline.glob(f'Workflow_{case}_0_*.txt'))
        ref = next((p for p in refs if p.name.endswith(f'_{repeat}.txt')), refs[-1])
        result = compare(ref, path)
        comparisons.append(dict(Case=case, Variant=variant, Repeat=repeat, ReferenceFile=ref.name, CandidateFile=path.name, **result))
    for label, directory in [('Baseline', baseline), ('Candidate', candidate)]:
        log = (directory / 'Native.txt').read_text(encoding='utf-8-sig')
        for line in log.splitlines():
            if not line.startswith('DATA_MTLI_WORKFLOW|'):
                continue
            parts = line.split('|')
            case, variant, repeat = map(int, parts[1:4])
            if repeat == 0:
                continue
            key = f'{label}_{case}_{variant}'
            timings.setdefault(key, []).append(float(parts[4].replace(',', '.')))
    summary = {key: dict(Samples=len(values), MedianSec=statistics.median(values), MinSec=min(values), MaxSec=max(values)) for key, values in timings.items()}
    for key, value in summary.items():
        label, case, variant = key.split('_')
        if label == 'Candidate':
            value['SpeedupAgainstRelease'] = summary[f'Baseline_{case}_0']['MedianSec'] / value['MedianSec']
    return dict(Reference='Unchanged release working VBA; same fixtures and geometry', Criterion='All meaningful numeric cells: absolute floor + 2% of reference; path changes listed separately', FullAcceptance=False, Timings=summary, Comparisons=comparisons)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('baseline', type=Path)
    parser.add_argument('candidate', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    result = analyse(args.baseline, args.candidate)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(result['Timings'], ensure_ascii=False, indent=2))
    print('Comparisons:', len(result['Comparisons']), 'outside 2%:', sum(not v['WithinTwoPercent'] for v in result['Comparisons']))
