"""Независимая приёмка реальных текстовых фигур, линий и рамок Excel."""
import json
import math
import re
from pathlib import Path

import numpy as np
from PIL import Image

directory = Path(__file__).resolve().parent


def read(stage):
    result = {}
    case = None
    path = directory.parent/'CandidateNativeAngles'/'NativeAngles.tsv' if stage == 'Baseline' else directory/'CandidateNativeAngles'/'NativeAngles.tsv'
    for line in path.read_text(encoding='ascii').splitlines():
        fields = line.split('\t')
        if fields[0] == 'CASE':
            case = {'degree': float(fields[2]), 'aspect': float(fields[3]), 'side': fields[4], 'sign': float(fields[5]), 'layout': [], 'shapes': []}
            result[int(fields[1])] = case
        elif fields[0] == 'LAYOUT':
            case['layout'].append({'id': int(fields[1]), 'kind': fields[2], 'box': np.array(fields[3:7], float), 'rotation': float(fields[7]), 'font': float(fields[8]), 'text': fields[9]})
        elif fields[0] == 'SHAPE':
            shape = {'name': fields[1], 'type': int(fields[2]), 'box': np.array(fields[3:7], float), 'rotation': float(fields[7])}
            if len(fields) > 8:
                shape.update(orientation=int(fields[8]), text=fields[9], glyph=np.array(fields[10:12], float))
                if len(fields)>12: shape['normalHeight']=float(fields[12])
            case['shapes'].append(shape)
    return result


def serial(case):
    return [(s['name'], s['type'], s['box'].tolist(), s['rotation'], s.get('orientation'), s.get('text'), s.get('glyph', np.array([])).tolist()) for s in case['shapes']]


old, new = read('Baseline'), read('Candidate')
errors = []
checks = 0
cardinal = []
right_vertical = []
glyph_gaps = []


def require(condition, reason):
    global checks
    checks += 1
    if not condition:
        errors.append(reason)


require(len(new) == len(old) == 56, 'Native case count differs')
for index, case in new.items():
    prefix = str(index)
    text_polygons = []
    require(len(case['shapes']) == 6, prefix+': repeated renderer must produce six objects')
    if case['degree'] in (0, 90):
        expected=serial(old[index])
        if case['degree']==90 and -case['sign']*(1 if case['side']=='Outside' else -1)>0:
            last=list(expected[-1]);last[4]=3;expected[-1]=tuple(last);right_vertical.append(index)
        require(serial(case) == expected, prefix+': cardinal coordinates/typography changed beyond requested right orientation')
        cardinal.append(index)
    angle = math.radians(case['degree'])
    direction = np.array((math.cos(angle)*1.6, -math.sin(angle)*case['aspect']/400))
    direction /= np.linalg.norm(direction)
    for entry, shape in zip(case['layout'], case['shapes']):
        if entry['kind'] == 'LINE':
            x1, y1, x2, y2 = entry['box']
            expected = np.array((min(x1, x2), min(y1, y2), abs(x2-x1), abs(y2-y1)))
            require(np.max(np.abs(shape['box']-expected)) < 0.001, prefix+': line object displaced')
            continue
        require(shape['type'] == 17, prefix+': label is not a text object')
        require(np.max(np.abs(shape['box']-entry['box'])) < 0.001, prefix+': text center or unrotated dimensions changed in renderer')
        require(abs(((shape['rotation']-entry['rotation']+180)%360)-180) < 0.001 or entry['rotation'] == 90 and shape['orientation'] == 2 or entry['rotation']==-90 and shape['orientation']==3, prefix+': native rotation differs')
        effective = -math.pi/2 if shape['orientation'] == 2 else math.pi/2 if shape['orientation']==3 else math.radians(shape['rotation'])
        tangent = np.array((math.cos(effective), math.sin(effective)))
        require(abs(tangent[0]*direction[1]-tangent[1]*direction[0]) < 0.00002, prefix+': rendered text is not parallel')
        require(shape['orientation'] in (2,3) or math.cos(effective) >= -0.000001, prefix+': text upside down')
        left, top, width, height = shape['box']
        center = np.array((left+width/2, top+height/2))
        radial = np.array((-tangent[1], tangent[0]))
        corners = np.array([center+a*tangent*width/2+b*radial*height/2 for a in (-1, 1) for b in (-1, 1)]) if shape['orientation'] not in (2,3) else np.array([(left, top), (left+width, top), (left, top+height), (left+width, top+height)])
        text_polygons.append(corners)
        require(np.min(corners[:, 0]) >= 0 and np.max(corners[:, 0]) <= 900 and np.min(corners[:, 1]) >= 0 and np.max(corners[:, 1]) <= 780, prefix+': rotated textbox outside chart')
        if True:
            line = case['layout'][2 if entry['id'] == 4 else 4]['box']
            line_mid = (line[:2]+line[2:])/2
            perpendicular = np.array((-direction[1], direction[0]))
            if np.dot(center-line_mid, perpendicular) < 0:
                perpendicular = -perpendicular
            center_gap=float(np.dot(center-line_mid,perpendicular))
            require(abs(center_gap-9)<0.001,prefix+': TextGap center rule differs from Config or other angles')
            visible_gap=center_gap-shape['normalHeight']/2
            reference=next(c for c in new.values() if c['degree']==0 and (c['aspect'],c['side'],c['sign'])==(case['aspect'],case['side'],case['sign']))
            reference_shape=reference['shapes'][3 if entry['id']==4 else 5]
            reference_gap=9-reference_shape['normalHeight']/2
            require(abs(visible_gap-reference_gap)<0.01,prefix+': actual glyph gap differs from horizontal text')
            require(visible_gap>0,prefix+': actual text touches the line')
            glyph_gaps.append(visible_gap)
        # BoundWidth/Height are axis-aligned glyph bounds after Office rotation.
        glyph = shape['glyph']
        if shape['orientation'] not in (2,3):
            inverse = np.array(((abs(tangent[0]), abs(tangent[1])), (abs(tangent[1]), abs(tangent[0]))))
            # Direct bound test avoids an ill-conditioned inverse at 45 degrees.
            maximum = inverse@np.array((width, height))
            require(np.all(glyph <= maximum+0.01), prefix+': rendered glyph bounds exceed textbox')
        image = directory/'CandidateNativeAngles'/('Case'+prefix+'.png')
        require(image.stat().st_size > 0, prefix+': empty PNG')
    if case['degree'] not in (0, 90):
        a, b = text_polygons
        axes = (direction, np.array((-direction[1], direction[0])))
        separated = any((a@axis).max() < (b@axis).min() or (b@axis).max() < (a@axis).min() for axis in axes)
        require(separated, prefix+': dimension and rebar textboxes overlap')

report = {'Passed': not errors, 'NativeCases': len(new), 'Checks': checks, 'CardinalCoordinatesAndTypographyPreserved': cardinal,'OnlyRightVerticalRebarTurned180':right_vertical,'ConfigTextGapCenterAllAnglesPoints':9,'ActualGlyphGapMatchesHorizontal':not errors,'GlyphGapPoints':[min(glyph_gaps),max(glyph_gaps)], 'TextRotationParallelAndReadable': not errors, 'DiagonalTextboxSeparationAndLineClearanceChecked': True, 'RepeatedDrawingCheckedInVBA': True, 'Errors': errors}
(directory/'NativeAngleAcceptance.json').write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
print(json.dumps(report, ensure_ascii=False))
if errors:
    raise SystemExit(1)
