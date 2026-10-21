"""Instrument the real Yes/No implementation, without experimental writer substitutions."""
from pathlib import Path
import json
import re

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]
SHARED = HERE.parent / 'PerformanceTradeoffs_2026-10-21'
OUT = HERE / 'ProductionModules'
RUN = HERE / 'ProductionBench02'
OUT.mkdir(exist_ok=True)
RUN.mkdir(exist_ok=True)
files = []

def replace(code, old, new):
    assert code.count(old) == 1, (old, code.count(old))
    return code.replace(old, new)

def save(name, code):
    path = OUT / name
    path.write_text(code, encoding='utf-8-sig', newline='\r\n')
    files.append(str(path.relative_to(ROOT)))

code = (ROOT / 'src/Excel/modWorkbookCalculation.bas').read_text(encoding='utf-8-sig')
code = replace(code, '    Set section = ReadFormedSectionFromResults(workbook)', '''    ProbeMark "calc.read.begin"
    Set section = ReadFormedSectionFromResults(workbook)
    Set ProbeCalculated = section
    ProbeMark "calc.prepare.begin"''')
code = replace(code, '    batch.Execute', '''    ProbeMark "calc.solve.begin"
    batch.Execute
    Set ProbeBatch = batch
    ProbeMark "calc.output.begin"''')
code = replace(code, '    ndmWriter.WriteResults workbook, section, PrepareSectionSnapshot(section, materialProvider), batch, units, True, fullFormatting', '    ndmWriter.WriteResults workbook, section, PrepareSectionSnapshot(section, materialProvider), batch, units, True, fullFormatting\n    ProbeMark "calc.plot.begin"')
code = replace(code, '    Dim totalElapsedSeconds As Double', '    ProbeMark "calc.finish.begin"\n    Dim totalElapsedSeconds As Double')
save('modWorkbookCalculation.bas', code)

# Generate the same scalar/element comparison from the current DTO source.
base = (SHARED / 'Build-Experiment.py').read_text(encoding='utf-8')
visitor = base[base.index('sources = {'):base.index("body = (HERE / 'Probe.bas.txt')")]
scope = {'ROOT': ROOT, 're': re}
exec(visitor, scope)
probe = (SHARED / 'Probe.bas.txt').read_text(encoding='utf-8')
probe = probe.replace('@ROOT@', str(RUN)).replace('@CASES@', '\n'.join(scope['cases']))
save('modSpeedTradeoffProbe.bas', probe)
runner = (HERE / 'Runner.bas.txt').read_text(encoding='utf-8')
runner = runner.replace('RunFormattingExperiment', 'RunFormattingFlagBenchmark')
runner = runner.replace('Array("Full", "Reuse", "FullEnd")', 'Array("Full", "Minimal", "FullEnd")')
runner = runner.replace('        ProbeVariant = "Full"\n        message', '        ProbeVariant = "Full"\n        SetSystemSetting "General.ReformatResultsEnabled", "Yes"\n        message')
runner = runner.replace('        ProbeVariant = CStr(choice)', '        ProbeVariant = "Full"\n        SetSystemSetting "General.ReformatResultsEnabled", IIf(CStr(choice) = "Minimal", "No", "Yes")')
runner = runner.replace('line = ProbeVariant & ","', 'line = CStr(choice) & ","')
runner = runner.replace('filename & "_" & ProbeVariant & ".xlsm"', 'filename & "_" & CStr(choice) & ".xlsm"')
save('modTestWorkbookInterface.bas', (ROOT / 'tests/modTestWorkbookInterface.bas').read_text(encoding='utf-8-sig') + '\n' + runner)
(HERE / 'ProductionModules.json').write_text(json.dumps(files, ensure_ascii=False, indent=2), encoding='utf-8')
print('Real Yes/No benchmark modules generated.')
