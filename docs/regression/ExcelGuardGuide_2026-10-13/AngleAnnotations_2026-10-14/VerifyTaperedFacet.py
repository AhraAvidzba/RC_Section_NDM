"""Проверяет фактический нативный пример: прежние стержни/размеры и две новые подписи."""
import json
import math
import struct
from pathlib import Path

directory = Path(__file__).resolve().parent
folder = directory / 'TaperedFacet02'

def fields(path, encoding):
    return [line.split('\t') for line in path.read_text(encoding=encoding).splitlines()]

def value(text):
    if text.startswith('d:'):
        return struct.unpack('>d', bytes.fromhex(text[2:]))[0]
    if text.startswith('i:'):
        return int(text[2:])
    return text[2:]

baseline = fields(directory/'CandidateGeometry/modTestShapeArchitecture.DumpShapeArchitectureSnapshot.txt', 'utf-8-sig')
before = fields(folder/'BeforeExtraLabel.tsv', 'utf-16')
after = fields(folder/'AfterExtraLabel.tsv', 'utf-16')
physical = {'CONCRETE','BAR','CONTOUR','FIBER','REBAR_SOURCE','PROPS','BOUNDS'}
expected = [line for line in baseline if len(line)>1 and line[0]=='RoundedTapered' and line[1] in physical]
actual = [line for line in after if len(line)>1 and line[1] in physical]
errors = []
if actual != expected: errors.append('Physical coordinates/properties differ from baseline')
old_annotations = [line for line in baseline if len(line)>1 and line[0]=='RoundedTapered' and line[1]=='ANNOTATION']
new_annotations = [line for line in after if len(line)>1 and line[1]=='ANNOTATION']
if not all(line in new_annotations for line in old_annotations): errors.append('An original annotation changed')
extra = [line for line in new_annotations if line not in old_annotations]
if len(extra)!=2: errors.append('Expected exactly two extra annotations')
label = next(line for line in extra if line[2]=='REBAR_UpperRightFacet')
if value(label[10])!='2Ø32': errors.append('Facet quantity/diameter differs from actual two bars')

snapshot = fields(folder/'GeometryAndShapes.tsv','utf-8-sig')
shapes = {line[2]: [value(v) for v in line[3:]] for line in snapshot if len(line)>2 and line[1]=='PLOT_SHAPE'}
text_shapes = [v for v in shapes.values() if v[0]==17 and v[-1] in ('2Ø32','L=0.201 m')]
if len(text_shapes)!=2: errors.append('Expected two actual native text shapes')
rotations = [v[5] for v in text_shapes]
if abs(rotations[0]-rotations[1])>0.001: errors.append('Native text angles differ')
dimension = next(v for name,v in shapes.items() if name=='NDMPlot_AnnotationLine387')
expected_angle = math.degrees(math.atan2(dimension[4], dimension[3]))
if max(abs(r-expected_angle) for r in rotations)>0.001: errors.append('Native text not parallel to actual dimension line')
native_run = json.loads((folder/'Run.json').read_text(encoding='utf-8-sig'))
result = {'Passed':not errors,'PhysicalSnapshotRowsUnchanged':len(actual),
          'OriginalAnnotationsPreserved':len(old_annotations),'AddedAnnotations':len(extra),
          'FacetLabel':'2Ø32','DimensionText':'L=0.201 m','NativeAngleDegrees':rotations[0],
          'DimensionLineAngleDegrees':expected_angle,'RepeatedDrawingIdentical':native_run['RepeatIdentical'],
          'Errors':errors}
(directory/'TaperedFacetAcceptance.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(result,ensure_ascii=False))
if errors: raise SystemExit(1)
