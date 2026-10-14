"""Сверяет реальные узлы Excel с независимой геометрией и обеими версиями схем."""
from pathlib import Path
from collections import defaultdict
from statistics import median
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import json
import math
import re

ROOT = Path(__file__).parent
TOLERANCE = 0.001  # пункта Excel; около 0,0013 пикселя при 96 DPI.


def number(value):
    return float(value.replace(',', '.'))


def read(folder):
    cases, times = {}, defaultdict(list)
    for text in (folder/'PlotObjects.txt').read_text(encoding='utf-8-sig').splitlines():
        row = text.split('|')
        if row[0] == 'TIME':
            times[row[1]].append(number(row[3]))
            continue
        case = cases.setdefault(int(row[1]), {'shapes': {}, 'nodes': defaultdict(list)})
        if row[0] == 'CASE': case['mode'], case['angle'], case['sweep'], case['size'] = row[2], number(row[3]), row[4], int(row[5])
        if row[0] in ('FRAME', 'DATA', 'ARC', 'COUNTS'): case[row[0].lower()] = list(map(number, row[2:]))
        if row[0] == 'SHAPE': case['shapes'][row[2]] = list(map(number, row[3:]))
        if row[0] == 'NODE': case['nodes'][row[2]].append(tuple(map(number, row[4:6])))
    return cases, times


def point(case, x, y):
    min_x, max_x, min_y, max_y, left, top, width, height = case['frame']
    return left+(x-min_x)/(max_x-min_x)*width, top+height-(y-min_y)/(max_y-min_y)*height


def clip(a, b, frame, inset=0.28):
    left, top, width, height = frame[4:]
    lo, hi = 0.0, 1.0
    for index, minimum, maximum in ((0, left+inset, left+width-inset), (1, top+inset, top+height-inset)):
        delta = b[index]-a[index]
        if abs(delta) < 1e-12:
            assert minimum <= a[index] <= maximum
            continue
        first, last = sorted(((minimum-a[index])/delta, (maximum-a[index])/delta))
        lo, hi = max(lo, first), min(hi, last)
    assert lo <= hi
    return tuple(a[i]+lo*(b[i]-a[i]) for i in (0, 1)), tuple(a[i]+hi*(b[i]-a[i]) for i in (0, 1))


def shared_shapes(case):
    grouped = defaultdict(list)
    for name, data in case['shapes'].items():
        if 'Principal' not in name and 'ContourArcLine' not in name:
            grouped[re.sub(r'\d+$', '', name)].append(data)
    return grouped


def main():
    old, old_times = read(ROOT/'BaselineObjects')
    new, new_times = read(ROOT/'CandidateObjects')
    errors, checks, node_checks, shared_checks = [], 0, 0, 0

    def close(a, b, reason):
        nonlocal checks
        checks += 1
        if len(a) != len(b) or any(abs(x-y) > TOLERANCE for x, y in zip(a, b)):
            errors.append(reason+': '+repr((a, b)))

    assert len(old) == len(new) == 162
    for index, current in new.items():
        previous = old[index]
        close(previous['frame'], current['frame'], f'Frame {index}')
        close(previous['data'], current['data'], f'Data {index}')
        a, b = shared_shapes(previous), shared_shapes(current)
        if a.keys() != b.keys(): errors.append(f'Other shape kinds {index}')
        for kind in a.keys() & b.keys():
            if len(a[kind]) != len(b[kind]): errors.append(f'Other shape count {index}: {kind}')
            for before, after in zip(a[kind], b[kind]):
                close(before, after, f'Other shape {index}: {kind}')
                shared_checks += 1
        expected_axes = 0 if current['mode'] == 'None' else 2
        if current['counts'][1] != expected_axes: errors.append(f'Axis object count {index}')
        if current['sweep'] == 'Circle':
            if current['counts'][2] != 0: errors.append(f'Circle turned into an arc {index}')
        else:
            if current['counts'][2] != 1: errors.append(f'Arc object count {index}')
            arc_name = next(name for name in current['nodes'] if 'ContourArcLine' in name)
            sx, sy, ex, ey, cx, cy, sweep = current['arc']
            radius = math.hypot(sx-cx, sy-cy)
            count = max(6, math.ceil(abs(sweep)/(math.pi/36)-1e-10))
            expected = [point(current, sx, sy)]
            start = math.atan2(sy-cy, sx-cx)
            for i in range(1, count):
                angle = start+sweep*i/count
                expected.append(point(current, cx+radius*math.cos(angle), cy+radius*math.sin(angle)))
            expected.append(point(current, ex, ey))
            actual = current['nodes'][arc_name]
            if len(actual) != len(expected): errors.append(f'Arc node count {index}')
            for i, (a, b) in enumerate(zip(actual, expected)):
                close(a, b, f'Arc node {index}:{i}')
                node_checks += 1
        if expected_axes:
            min_x, max_x, min_y, max_y, cx, cy, angle = current['data']
            length = 0.65*math.hypot(max_x-min_x, max_y-min_y)
            for axis, direction in ((1, (math.cos(angle), math.sin(angle))), (2, (-math.sin(angle), math.cos(angle)))):
                a = point(current, cx-length*direction[0], cy-length*direction[1])
                b = point(current, cx+length*direction[0], cy+length*direction[1])
                a, b = clip(a, b, current['frame'])
                name = next(name for name in current['nodes'] if f'Principal{axis}' in name)
                nodes = current['nodes'][name]
                if len(nodes) != 4: errors.append(f'Axis band node count {index}:{axis}')
                first = tuple((nodes[0][i]+nodes[3][i])/2 for i in (0, 1))
                last = tuple((nodes[1][i]+nodes[2][i])/2 for i in (0, 1))
                close(first, a, f'Axis start {index}:{axis}')
                close(last, b, f'Axis end {index}:{axis}')
        if current['size'] == 1:
            unmoved = new[index-1]
            close(unmoved['frame'], current['frame'], f'Chart move frame {index}')
            for name, nodes in current['nodes'].items():
                prior = unmoved['nodes'].get(name)
                if prior is None: errors.append(f'Chart move object {index}: {name}')
                else:
                    for a, b in zip(nodes, prior): close(a, b, f'Chart move nodes {index}: {name}')
    arc_paths = 0
    for path in (ROOT/'CandidateObjects').glob('Case*.xlsx'):
        with ZipFile(path) as archive:
            for name in archive.namelist():
                if name.startswith('xl/drawings/') and name.endswith('.xml'):
                    document = ET.fromstring(archive.read(name))
                    for shape in document.iter():
                        if shape.tag.rsplit('}', 1)[-1] != 'sp': continue
                        if not any('ContourArcLine' in child.get('name', '') for child in shape.iter()): continue
                        arc_paths += 1
                        if any(child.tag.rsplit('}', 1)[-1] == 'close' for child in shape.iter()): errors.append('Unexpected arc closure: '+path.name)
    assert arc_paths == 6
    times = {mode: {'BeforeMedian': median(old_times[mode]), 'AfterMedian': median(new_times[mode])} for mode in old_times}
    result = {'Passed': not errors, 'Cases': len(new), 'Checks': checks, 'ArcNodesChecked': node_checks,
              'OtherShapesCompared': shared_checks, 'UnclosedArcPathsInSavedBooks': arc_paths,
              'AxesObjects': '2 when enabled, 0 when disabled', 'ArcObjects': '1 per arc',
              'DrawTimesSeconds': times,
              'ObjectCounts': {'Before': [min(c['counts'][0] for c in old.values()), max(c['counts'][0] for c in old.values())],
                               'After': [min(c['counts'][0] for c in new.values()), max(c['counts'][0] for c in new.values())]},
              'Errors': errors}
    (ROOT/'NativeGeometryAcceptance.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False))
    if errors: raise SystemExit(1)


if __name__ == '__main__': main()
