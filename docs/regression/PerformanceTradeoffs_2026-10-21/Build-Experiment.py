"""Build isolated VBA instrumentation; never modify production sources or book."""
from pathlib import Path
import re
import json

ROOT = Path(__file__).resolve().parents[3]
HERE = Path(__file__).resolve().parent
OUT = HERE / 'Modules'
OUT.mkdir(exist_ok=True)
files = []

def read(name):
    return (ROOT / name).read_text(encoding='utf-8-sig')

def replace(text, old, new, count=1):
    assert text.count(old) == count, (old, text.count(old), count)
    return text.replace(old, new)

def save(name, text):
    p = OUT / Path(name).name
    p.write_text(text, encoding='utf-8-sig', newline='\r\n')
    files.append(str(p.relative_to(ROOT)))

def entry(text, method, code):
    pattern = r'(?m)^((?:Public|Private|Friend) (?:Sub|Function) ' + method + r'\([^\n]*(?:\n[^\n]*)*?\n)'
    # Signature may have continuation lines; stop after first non-continuation.
    lines = text.splitlines(True)
    index = next(i for i, line in enumerate(lines) if re.match(r'(Public|Private|Friend) (Sub|Function) '+method+r'\(', line))
    while lines[index].rstrip().endswith('_'):
        index += 1
    lines.insert(index+1, code + '\n')
    return ''.join(lines)

name = 'src/Excel/modWorkbookCalculation.bas'
s = read(name)
s = replace(s, '    guard.Enter workbook.Application\n    Dim settings', '    ProbeMark "form.guard.begin"\n    guard.Enter workbook.Application\n    ProbeMark "form.inputs.begin"\n    Dim settings')
s = replace(s, '        Set section = BuildWorkbookSectionModel(workbook, settings, units)', '        ProbeMark "form.model.begin"\n        Set section = BuildWorkbookSectionModel(workbook, settings, units)')
s = replace(s, '    Dim snapshot As CSectionPropertiesCalculator: Set snapshot = PrepareSectionSnapshot(section)', '    Set ProbeFormed = section\n    ProbeMark "form.properties.begin"\n    Dim snapshot As CSectionPropertiesCalculator: Set snapshot = PrepareSectionSnapshot(section)')
s = replace(s, '    summary.ClearSummary workbook\n    writer.WriteGeometryPreview workbook, section, snapshot, units', '    ProbeMark "form.output.begin"\n    summary.ClearSummary workbook\n    If ProbeMemory Then\n        writer.ClearResults workbook\n    Else\n        writer.WriteGeometryPreview workbook, section, snapshot, units\n    End If\n    ProbeMark "form.plot.begin"')
s = replace(s, '    guard.Restore True\n    FormSectionForWorkbook', '    ProbeMark "form.restore.begin"\n    guard.Restore True\n    ProbeMark "form.end"\n    FormSectionForWorkbook')
s = replace(s, '    excelGuard.Enter workbook.Application', '    ProbeMark "calc.guard.begin"\n    excelGuard.Enter workbook.Application\n    ProbeMark "calc.inputs.begin"')
s = replace(s, '    Set section = ReadFormedSectionFromResults(workbook)', '    ProbeMark "calc.read.begin"\n    If ProbeMemory Then\n        Set section = ProbeFormed\n    Else\n        Set section = ReadFormedSectionFromResults(workbook)\n    End If\n    Set ProbeCalculated = section\n    ProbeMark "calc.prepare.begin"')
s = replace(s, '    batch.Execute', '    ProbeMark "calc.solve.begin"\n    batch.Execute\n    Set ProbeBatch = batch\n    ProbeMark "calc.output.begin"')
s = replace(s, '    ndmWriter.WriteResults workbook, section, PrepareSectionSnapshot(section, materialProvider), batch, units, True', '    If Not ProbeMemory Then ndmWriter.WriteResults workbook, section, PrepareSectionSnapshot(section, materialProvider), batch, units, True\n    ProbeMark "calc.plot.begin"')
s = replace(s, '    Dim totalElapsedSeconds As Double', '    ProbeMark "calc.finish.begin"\n    Dim totalElapsedSeconds As Double')
s = replace(s, '    excelGuard.Restore True\n    summaryWriter.UpdateElapsedSeconds', '    excelGuard.Restore True\n    ProbeMark "calc.end"\n    summaryWriter.UpdateElapsedSeconds')
save(name, s)

name = 'src/Excel/CNDMResultsWriter.cls'
s = read(name)
s = entry(s, 'FormatResultsBlock', '    If ProbePlain Then Exit Sub')
s = replace(s, '    WriteElementResults workbook, section, batch, runID, units', '    If Not ProbeMinimum Then WriteElementResults workbook, section, batch, runID, units')
s = replace(s, '    WriteMaterialDiagrams workbook, batch, runID, units', '    If Not ProbeMinimum Then WriteMaterialDiagrams workbook, batch, runID, units')
s = replace(s, '    WriteSectionAnnotations workbook, section, runID, units, batch, previousAnnotations', '    If Not ProbeMinimum Then WriteSectionAnnotations workbook, section, runID, units, batch, previousAnnotations')
s = replace(s, '    WriteSectionAnnotations workbook, section, runID, units\n', '    If Not ProbeMinimum Then WriteSectionAnnotations workbook, section, runID, units\n')
s = s.replace('data(rowIndex, 16) = SectionOptionalJson', 'If Not ProbeNoJson Then data(rowIndex, 16) = SectionOptionalJson')
s = s.replace('data(rowIndex, 24) = SectionOptionalJson', 'If Not ProbeNoJson Then data(rowIndex, 24) = SectionOptionalJson')
for assignment in [
    'target.Resize(rowCount, GEOMETRY_COLUMNS).Value2 = data',
    'target.Resize(contours.Count + 1, CONTOUR_COLUMNS).Value2 = data',
    'target.Resize(rowCount, MATERIAL_DIAGRAM_COLUMNS).Value2 = data',
    'target.Resize(rowCount, ANNOTATION_COLUMNS).Value2 = data',
    'target.Resize(1, ANNOTATION_COLUMNS).Value2 = data',
]:
    s = replace(s, '    ' + assignment, '    If ProbeJsonEnabled Then\n        ProbeWriteJson target, data\n        Exit Sub\n    End If\n    ' + assignment)
s = entry(s, 'UpdateGeometrySnapshotRunID', '    If ProbeJsonEnabled Then ProbeJsonRunId workbook, runID: Exit Sub')
save(name, s)

name = 'src/Excel/CBatchResultWriter.cls'
s = read(name)
s = entry(s, 'FormatSummary', '    If ProbePlain Then Exit Sub')
s = replace(s, '    stage = "WriteStrengthSummary"', '    If ProbeMinimum Then Exit Sub\n    stage = "WriteStrengthSummary"')
save(name, s)
for cls in ['CStrengthSummaryWriter', 'CCrackSummaryWriter', 'CStabilitySummaryWriter']:
    name = 'src/Excel/' + cls + '.cls'
    save(name, entry(read(name), 'FormatTable', '    If ProbePlain Then Exit Sub'))

name = 'src/Excel/CSectionPlotter.cls'
s = read(name)
s = entry(s, 'ShouldDrawElementShapeLayer', '    If ProbeVariant = "Grouped" Then Exit Function\n    If ProbeVariant = "Outline" Then ShouldDrawElementShapeLayer = True: Exit Function')
s = replace(s, '            Set shapeObject = AddConcreteElementShape(data, itemIndex, fallbackSquareRotation, colorValue)', '            If ProbeVariant <> "Outline" Then Set shapeObject = AddConcreteElementShape(data, itemIndex, fallbackSquareRotation, colorValue)')
s = entry(s, 'DrawLegend', '    If ProbeVariant = "Outline" Then Exit Sub')
save(name, s)

save('tests/modTestWorkbookInterface.bas', read('tests/modTestWorkbookInterface.bas') + '\n' + (HERE / 'Runner.bas.txt').read_text(encoding='utf-8'))

# Generate comparisons from actual scalar properties in result DTOs, not cell output.
sources = {p.stem: p.read_text(encoding='utf-8-sig') for p in (ROOT / 'src').rglob('*.cls')}
visited = set()
cases = []
def visit(cls):
    if cls in visited: return
    visited.add(cls)
    code = ['        Case "' + cls + '"']
    for prop, typ in re.findall(r'^Public Property Get (\w+)\(\) As (\w+)', sources[cls], re.M):
        if prop in ['DiagnosticLog', 'ElapsedSeconds', 'ElapsedSecondsText']: continue
        if typ in ['String', 'Boolean', 'Double', 'Long', 'Integer', 'Single', 'Currency'] or typ.startswith('E'):
            code.append(f'            ProbeValue prefix & ".{prop}", CallByName(obj, "{prop}", VbGet)')
        elif typ in sources and (typ.endswith('Result') or typ == 'CResultMeta'):
            code.append(f'            Set child = CallByName(obj, "{prop}", VbGet)')
            code.append(f'            ProbeObject child, "{typ}", prefix & ".{prop}"')
            visit(typ)
    cases.append('\n'.join(code))
visit('CCombinationResult')
visit('CSectionStateResult')
body = (HERE / 'Probe.bas.txt').read_text(encoding='utf-8')
body = body.replace('@ROOT@', str(HERE)).replace('@CASES@', '\n'.join(cases))
save('modSpeedTradeoffProbe.bas', body)
save('modCompactJsonProbe.bas', (HERE / 'JsonProbe.bas.txt').read_text(encoding='utf-8'))
for name in ['src/Excel/modSectionSnapshot.bas', 'src/Excel/CSectionPlotDataReader.cls']:
    s = read(name)
    start = s.index('    Dim columnCount As Long\n', s.index('Function ReadAnchoredResultTable('))
    end = s.index('    ReadAnchoredResultTable = anchor.Resize(rowCount, columnCount).Value2', start)
    end += len('    ReadAnchoredResultTable = anchor.Resize(rowCount, columnCount).Value2')
    s = s[:start] + '    If ProbeJsonRange(rangeName) Then\n        ReadAnchoredResultTable = ProbeReadJson(workbook, rangeName)\n    Else\n' + s[start:end] + '\n    End If' + s[end:]
    if name.endswith('modSectionSnapshot.bas'):
        s = replace(s, '    If CStr(geometryAnchor.Value2) = "RunID" Then snapshotRun = geometryAnchor.Offset(1, 0).Value2', '    If ProbeJsonEnabled Then\n        snapshotRun = geometryAnchor.Offset(0, 4).Value2\n    Else\n        If CStr(geometryAnchor.Value2) = "RunID" Then snapshotRun = geometryAnchor.Offset(1, 0).Value2\n    End If')
    save(name, s)
(HERE / 'Modules.json').write_text(json.dumps(files, ensure_ascii=False, indent=2), encoding='utf-8')
print('Generated', len(files), 'diagnostic modules; production untouched.')
