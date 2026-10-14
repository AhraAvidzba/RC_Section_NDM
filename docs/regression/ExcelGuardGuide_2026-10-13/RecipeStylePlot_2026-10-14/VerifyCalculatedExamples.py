"""Независимо связывает фактические Excel-объекты/PNG с рассчитанной геометрией."""
import json
import re
from pathlib import Path

import numpy as np
from PIL import Image

directory = Path(__file__).resolve().parent
errors = []
cases = []


def require(value, reason):
    if not value:
        errors.append(reason)


def read_data(folder):
    rows = [line.split('\t') for line in (folder/'CalculatedData.tsv').read_text(encoding='utf-8-sig').splitlines() if line]
    return {
        'state': next(r[1:] for r in rows if r[0] == 'STATE'),
        'centers': np.array(next(r[1:] for r in rows if r[0] == 'CENTERS'), float),
        'elements': [(r[1], r[2], np.array(r[3:], float)) for r in rows if r[0] == 'ELEMENT'],
        'contours': [(r[1], np.array(r[2:], float)) for r in rows if r[0] == 'ANNOTATION' and r[1] == 'CONTOUR_LINE'],
    }


def endpoints(shape):
    n = np.array(shape['Nodes'], float)
    return (n[0]+n[3])/2, (n[1]+n[2])/2


def cross_point(axis1, axis2):
    a, b = endpoints(axis1)
    c, d = endpoints(axis2)
    t = np.linalg.solve(np.column_stack((b-a, -(d-c))), c-a)[0]
    return a+t*(b-a)


base = read_data(directory/'CalculatedBaselineG')
new = read_data(directory/'CalculatedGSection')
require(base['state'] == new['state'], 'G: calculation status changed')
require(np.array_equal(base['centers'], new['centers']), 'G: saved calculation centers changed')
require(len(base['elements']) == len(new['elements']) and all(a[:2] == b[:2] and np.array_equal(a[2], b[2]) for a, b in zip(base['elements'], new['elements'])), 'G: geometry or calculated element stresses changed')
require(base['contours'] and len(base['contours']) == len(new['contours']) and all(a[0] == b[0] and np.array_equal(a[1], b[1]) for a, b in zip(base['contours'], new['contours'])), 'G: contour changed')

for name in ('GSection', 'HollowRounded'):
    folder = directory/('Calculated'+name)
    data = read_data(folder)
    require(data['state'][0] == 'OK' and data['state'][1] == 'False', name+': not a converged admissible calculation')
    concrete = np.array([v for _, kind, v in data['elements'] if kind == 'Concrete'])
    centroid = np.sum(concrete[:, :2]*concrete[:, 2:3], axis=0)/np.sum(concrete[:, 2])
    require(np.max(np.abs(centroid-data['centers'][:2])) < 1e-8, name+': concrete centroid differs from independent area sum')
    require(np.max(np.abs(centroid-data['centers'][6:8])) < 1e-8, name+': zero load offsets do not give concrete centroid')
    original = {}
    for file in sorted(folder.glob('*Original.json'))+sorted(folder.glob('*Moved.json'))+sorted(folder.glob('*Resized.json')):
        case = json.loads(file.read_text(encoding='utf-8-sig'))
        label = name+'/'+case['Mode']+'/'+case['Frame']
        shapes = case['Shapes']
        require(shapes == case['BeforeExport'], label+': activation/export moved native shapes')
        elements = [s for s in shapes if s['Name'].startswith('NDMPlot_ElementConcrete')]
        require(len(elements) == len(concrete), label+': concrete object count mismatch')
        points = np.array([[s['Left']+s['Width']/2, s['Top']+s['Height']/2] for s in elements])
        # Affine map is inferred from every actual concrete object, without plotter internals.
        fit = np.linalg.lstsq(np.column_stack((concrete[:, :2], np.ones(len(concrete)))), points, rcond=None)[0]
        residual = np.max(np.abs(np.column_stack((concrete[:, :2], np.ones(len(concrete))))@fit-points))
        require(residual < 0.001, label+': geometry coordinates do not share a single affine map')
        transform = lambda point: np.append(point, 1)@fit
        axes = [s for s in shapes if re.match(r'NDMPlot_Principal[12]\d+$', s['Name'])]
        require(len(axes) == 2 and all(s['Type'] == 5 for s in axes), label+': axes must be two Freeforms')
        intersection = cross_point(*axes)
        index = 0 if case['Mode'] == 'Concrete' else 2
        expected = transform(data['centers'][index:index+2])
        axis_error = float(np.max(np.abs(intersection-expected)))
        require(axis_error < 0.001, label+': axes intersection is displaced relative to concrete geometry')
        marker = next(s for s in shapes if s['Name'] == 'NDMPlot_LoadPoint')
        marker_center = np.array((marker['Left']+marker['Width']/2, marker['Top']+marker['Height']/2))
        marker_error = float(np.max(np.abs(marker_center-transform(centroid))))
        require(marker_error < 0.001, label+': load marker is displaced relative to concrete geometry')
        if case['Mode'] == 'Concrete':
            require(np.max(np.abs(intersection-marker_center)) < 0.001, label+': concrete axes and zero-offset load point must coincide')
        else:
            require(np.max(np.abs((intersection-marker_center)-(transform(data['centers'][2:4])-transform(centroid)))) < 0.001, label+': transformed axes/load relative displacement is wrong')
        # Compare outer contour bounds to the same map independently fitted to concrete.
        lines = [s for s in shapes if re.fullmatch(r'NDMPlot_Contour\d+', s['Name'])]
        require(len(lines) == len(data['contours']), label+': contour line count differs')
        for shape, (_, coords) in zip(lines, data['contours']):
            a, b = transform(coords[:2]), transform(coords[2:4])
            expected_bounds = np.array((min(a[0], b[0]), min(a[1], b[1]), abs(b[0]-a[0]), abs(b[1]-a[1])))
            require(np.max(np.abs(expected_bounds-np.array([shape[k] for k in ('Left', 'Top', 'Width', 'Height')]))) < 0.001, label+': contour is displaced')
        image = np.array(Image.open(file.with_suffix('.png')).convert('RGB'))
        scale = np.array((image.shape[1]/case['ChartWidth'], image.shape[0]/case['ChartHeight']))
        # Excel exports Chart.Shapes with a common raster inset. Establish it from
        # the thick outer/inner contour, independently of both axes and load marker.
        offsets = [[], []]
        dark = np.max(image, axis=2) < 100
        for line in lines:
            for dimension in (0, 1):
                if line[('Width', 'Height')[dimension]] > 0.01:
                    continue
                fixed = line[('Left', 'Top')[dimension]]*scale[dimension]
                start = line[('Top', 'Left')[dimension]]*scale[1-dimension]
                length = line[('Height', 'Width')[dimension]]*scale[1-dimension]
                if length < 25:
                    continue
                lo, hi = int(start+length*0.15), int(start+length*0.85)
                candidates = range(max(0, int(fixed)-12), min(image.shape[1-dimension], int(fixed)+13))
                scores = [int(np.sum(dark[lo:hi, p] if dimension == 0 else dark[p, lo:hi])) for p in candidates]
                best = max(scores)
                require(best > (hi-lo)*0.7, label+': native contour not found in PNG')
                positions = [p for p, score in zip(candidates, scores) if score >= best-1]
                offsets[dimension].append(float(np.mean(positions))-fixed)
        inset = np.array([np.median(v) for v in offsets])
        require(all(len(v) >= 2 for v in offsets), label+': insufficient raster contour calibration')
        orange = (image[:, :, 0] > 235) & (image[:, :, 1] > 95) & (image[:, :, 1] < 160) & (image[:, :, 2] < 40)
        yy, xx = np.where(orange)
        require(len(xx) > 5, label+': orange load marker not visible in actual PNG')
        png_marker = np.array(((xx.min()+xx.max())/2, (yy.min()+yy.max())/2))
        raster_error = float(np.max(np.abs(png_marker-(marker_center*scale+inset))))
        require(raster_error < 1.5, label+': PNG marker shifted from native objects')
        # Sample the rendered black axes near the crossing, avoiding the diamond itself.
        rendered = []
        for axis in axes:
            a, b = endpoints(axis); direction = (b-a)/np.linalg.norm(b-a)
            for distance in (-30, -20, 20, 30):
                point = (intersection+distance*direction)*scale+inset
                x, y = np.rint(point).astype(int)
                pixels = image[max(0, y-2):y+3, max(0, x-2):x+3]
                rendered.append(bool(np.any(np.max(pixels, axis=2) < 100)))
        require(all(rendered), label+': actual PNG axes not found at geometric position')
        if case['Frame'] == 'Original':
            original[case['Mode']] = points
        elif case['Frame'] == 'Moved':
            require(np.max(np.abs(original[case['Mode']]-points)) < 0.001, label+': moving chart displaced relative geometry')
        cases.append({'Case': label, 'ConcreteObjects': len(elements), 'AxesIntersectionErrorPt': axis_error, 'LoadMarkerErrorPt': marker_error, 'PNGMarkerErrorPx': raster_error, 'RasterContourInsetPx': inset.tolist(), 'RenderedAxisSamplesFound': sum(rendered), 'IntersectionPt': intersection.tolist(), 'LoadMarkerPt': marker_center.tolist()})

report = {'Passed': not errors, 'BaselineGCalculationUnchanged': True, 'Cases': cases, 'Errors': errors}
(directory/'CalculatedExamplesAcceptance.json').write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
print(json.dumps(report, ensure_ascii=False))
if errors:
    raise SystemExit(1)
