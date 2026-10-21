"""Profile repeated Results formatting in isolated copies only."""
from pathlib import Path
import json
import subprocess

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]
SHARED = HERE.parent / 'PerformanceTradeoffs_2026-10-21'

# Reuse the measured full-command timer and DTO comparison, without touching
# the scripts or evidence of the previous experiment.
base = (SHARED / 'Build-Experiment.py').read_text(encoding='utf-8')
base = base.replace("return (ROOT / name).read_text(encoding='utf-8-sig')", "return subprocess.check_output(['git','show','33181b12ab9b8bcb49a72f22fab3a0905819ca61:'+name]).decode('utf-8-sig')")
base = 'import subprocess\n' + base
base = base.replace("(HERE / 'Runner.bas.txt')", "(SHARED / 'Runner.bas.txt')")
base = base.replace("(HERE / 'Probe.bas.txt')", "(SHARED / 'Probe.bas.txt')")
base = base.replace("(HERE / 'JsonProbe.bas.txt')", "(SHARED / 'JsonProbe.bas.txt')")
scope = {'__file__': str(HERE / 'Build-Experiment.py'), 'SHARED': SHARED}
exec(compile(base, str(SHARED / 'Build-Experiment.py'), 'exec'), scope)
OUT = scope['OUT']
entry = scope['entry']
replace = scope['replace']
files = scope['files']
RUN = HERE / 'Run02'
RUN.mkdir(exist_ok=True)

def modify(name, transform):
    path = OUT / name
    code = transform(path.read_text(encoding='utf-8-sig'))
    path.write_text(code, encoding='utf-8-sig', newline='\r\n')

def detailed(code, status_call):
    code = entry(code, 'PrepareTarget', '''    If FormattingReuse Then
        If target.Rows.Count > HEADER_ROWS Then target.Offset(HEADER_ROWS, 0).Resize(target.Rows.Count - HEADER_ROWS, TOTAL_COLUMNS).ClearContents
        Exit Sub
    End If''')
    code = entry(code, 'WriteBlockTitle', '    If FormattingReuse Then target.Worksheet.Cells(target.Row - 1, target.Column).Value2 = titleText: Exit Sub')
    code = entry(code, 'WriteHeaders', '    If FormattingReuse Then Exit Sub')
    code = entry(code, 'FormatTable', f'''    If FormattingReuse Then
        If FormattingStatuses Then {status_call}
        Exit Sub
    End If''')
    return code

modify('CStrengthSummaryWriter.cls', lambda s: detailed(s, 'ApplyStatusCells target, batch, target.Rows.Count - HEADER_ROWS'))
modify('CCrackSummaryWriter.cls', lambda s: detailed(s, 'ApplyStatusCells target, batch, target.Rows.Count - HEADER_ROWS'))
modify('CStabilitySummaryWriter.cls', lambda s: detailed(s, 'FormatStatusCells target, batch'))

def compact(code):
    code = entry(code, 'ClearSummaryBody', '''    If FormattingReuse Then
        target.Cells(1, 1).Resize(META_ROWS, 1).ClearContents
        target.Cells(1, 7).Resize(3, 1).ClearContents
        target.Offset(WORST_HEADER_ROW - 1, 0).Resize(target.Rows.Count - WORST_HEADER_ROW + 1, TARGET_COLUMNS).ClearContents
        Exit Sub
    End If''')
    code = entry(code, 'WriteSummaryBody', '''    If FormattingReuse Then
        FormattingWriteSlice target, data, 1, 1, META_ROWS, 1
        FormattingWriteSlice target, data, 1, 7, 3, 1
        FormattingWriteSlice target, data, WORST_HEADER_ROW, 1, UBound(data, 1) - WORST_HEADER_ROW + 1, TARGET_COLUMNS
        Exit Sub
    End If''')
    code = entry(code, 'FormatSummary', '''    If FormattingReuse Then
        If FormattingStatuses Then FormatStatusCells target, batch
        Exit Sub
    End If''')
    return code

modify('CBatchResultWriter.cls', compact)

def snapshots(code):
    code = entry(code, 'FormatResultsBlock', '''    If FormattingReuse Then
        If rowCount > 1 Then target.Offset(1, 0).Resize(rowCount - 1, columnCount).WrapText = False
        Exit Sub
    End If''')
    code = entry(code, 'WriteSnapshotTitle', '    If FormattingReuse Then Exit Sub')
    code = entry(code, 'ClearNamedRegion', '''    If FormattingReuse Then
        Dim existingAnchor As Object, existingRows As Long
        Set existingAnchor = workbook.Names.Item(rangeName).RefersToRange
        existingRows = ExistingAnchoredRowCount(existingAnchor)
        If existingRows > 1 Then existingAnchor.Offset(1, 0).Resize(existingRows - 1, OutputColumnCount(rangeName)).ClearContents
        Exit Sub
    End If''')
    # Mandatory text columns already have their formats in the prepared output.
    code = replace(code, '    target.Offset(0, 14).Resize(rowCount, 2).NumberFormat = "@"', '    If Not FormattingReuse Then target.Offset(0, 14).Resize(rowCount, 2).NumberFormat = "@"')
    code = replace(code, '        target.Resize(rowCount, ANNOTATION_COLUMNS).Columns.Item(CLng(textColumn)).NumberFormat = "@"', '        If Not FormattingReuse Then target.Resize(rowCount, ANNOTATION_COLUMNS).Columns.Item(CLng(textColumn)).NumberFormat = "@"')
    return code

modify('CNDMResultsWriter.cls', snapshots)
modify('modSpeedTradeoffProbe.bas', lambda s: s.replace(str(HERE), str(RUN)))
modify('modTestWorkbookInterface.bas', lambda s: s + '\n' + (HERE / 'Runner.bas.txt').read_text(encoding='utf-8'))
path = OUT / 'modResultsFormattingProbe.bas'
path.write_text((HERE / 'Probe.bas.txt').read_text(encoding='utf-8'), encoding='utf-8-sig', newline='\r\n')
files.append(str(path.relative_to(ROOT)))
(HERE / 'Modules.json').write_text(json.dumps(files, ensure_ascii=False, indent=2), encoding='utf-8')
print('Formatting probe ready; production untouched.')
