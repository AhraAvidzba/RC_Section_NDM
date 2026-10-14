"""Приёмка книги: данные прежние, изменены только шесть компонентов и схема."""
import argparse
import hashlib
import importlib.util
import json
import re
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

root = Path(__file__).resolve().parents[4]
spec = importlib.util.spec_from_file_location('presentation', root/'tools/build_workbook/Verify-GuidePresentation.py')
presentation = importlib.util.module_from_spec(spec)
spec.loader.exec_module(presentation)
reader = presentation.reader

parser = argparse.ArgumentParser()
parser.add_argument('--directory', type=Path, default=Path(__file__).resolve().parent)
parser.add_argument('--book', type=Path)
parser.add_argument('--dump', type=Path)
parser.add_argument('--report', type=Path)
args = parser.parse_args()
directory = args.directory
book, dump = args.book or directory/'Candidate.xlsm', args.dump or directory/'FinalNativeVBA.txt'
old, new = reader.read(directory/'Baseline.xlsm'), reader.read(book)
errors = []
checked = 0


def require(value, reason):
    if not value:
        errors.append(reason)


require(old['names'] == new['names'], 'Named ranges changed')
require(old['sheets'].keys() == new['sheets'].keys(), 'Sheet inventory changed')
for sheet, data in old['sheets'].items():
    actual = new['sheets'][sheet]
    for address in data['cells'].keys() | actual['cells'].keys():
        checked += 1
        require(data['cells'].get(address) == actual['cells'].get(address), 'Cell changed: '+sheet+'!'+address)
    require(data['rows'] == actual['rows'], 'Row heights/styles/visibility changed: '+sheet)
    for tag in {e.tag for e in data['document']} | {e.tag for e in actual['document']}:
        if tag in (reader.S+'sheetData', reader.S+'sheetViews'):
            continue
        a, b = data['document'].find(tag), actual['document'].find(tag)
        if tag == reader.S+'mergeCells':
            require({e.get('ref') for e in a} == {e.get('ref') for e in b}, 'Merges changed: '+sheet)
        else:
            require(reader.canonical(a) == reader.canonical(b), 'Sheet structure changed: '+sheet+':'+tag)

require(old['drawings'].keys() == new['drawings'].keys(), 'Drawing inventory changed')
changed_drawings = []
for name, value in old['drawings'].items():
    if value != new['drawings'][name]:
        changed_drawings.append(name)
        require(name == 'xl/drawings/drawing4.xml', 'Drawing outside section plot changed: '+name)
baseline, final = presentation.components(directory/'BaselineNativeVBA.txt'), presentation.components(dump)
allowed = {'CSectionRecipeCircle', 'CSectionRecipeRectSet', 'CSectionRecipeRoundedRectangle', 'CSectionRecipeHollowRectangle', 'CSectionPlotter', 'modTestPlotConfig'}
require(baseline.keys() == final.keys() and len(final) == 157, 'Native component inventory changed')
def canonical_native(component):
    # VBE changes identifier case across a project when a module is imported.
    # Preserve comments, string literals, numbers and whitespace exactly.
    pattern = r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27\n]+|\n'
    body = ''.join(part if part.startswith(('"', "'")) else part.lower() for part in re.findall(pattern, component[2]))
    return component[:2] + (body,)

raw_changed = {name for name in baseline if baseline[name] != final.get(name)}
changed = {name for name in baseline if canonical_native(baseline[name]) != canonical_native(final[name])}
require(changed == allowed, 'Unexpected native component changes: '+repr(changed))
require(all('GoalCalculatedExample' not in body[2] and 'GoalCapturedFrame' not in body[2] and 'GoalPlotObjectProbe' not in body[2] for body in final.values()), 'Temporary test instrumentation leaked into workbook')
markers = presentation.embedded_markers(book, dump)
require(not markers, 'Embedded technical markers: '+repr(markers))
with ZipFile(book) as archive:
    xml = ET.fromstring(archive.read('xl/drawings/drawing4.xml'))
    names = [e.get('name') for e in xml.iter() if e.tag.endswith('}cNvPr')]
    axes = [n for n in names if re.fullmatch(r'NDMPlot_Principal[12]\d+', n or '')]
    require(len(axes) == 2, 'Actual saved plot must have exactly two principal axes')
    require(not any('Dash' in (n or '') for n in names), 'Saved plot retains fragmented axes')
    arc_names = [n for n in names if (n or '').startswith('NDMPlot_ContourArcLine')]
    require(len(arc_names) == len(set(arc_names)), 'Arc object names duplicated')

sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
report = {'Passed': not errors, 'CellsChecked': checked, 'NamedRanges': len(old['names']), 'InputResultsFormulasStylesAndHelpUnchanged': not any(e.startswith(('Cell changed:', 'Row ', 'Sheet ', 'Named ')) for e in errors), 'ChangedDrawings': changed_drawings, 'NativeComponents': len(final), 'ChangedNativeComponents': sorted(changed), 'VBEIdentifierCaseOnlyComponents': sorted(raw_changed-changed), 'NativeChangesRestricted': changed == allowed, 'TemporaryHelpersAbsent': True, 'NoEmbeddedMarkers': not markers, 'SavedPrincipalAxesObjects': len(axes), 'SavedContourArcObjects': len(arc_names), 'CandidateSHA256': sha(book), 'NativeSHA256': sha(dump), 'Errors': errors}
(args.report or directory/'Preservation.json').write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
print(json.dumps(report, ensure_ascii=False))
if errors:
    raise SystemExit(1)
