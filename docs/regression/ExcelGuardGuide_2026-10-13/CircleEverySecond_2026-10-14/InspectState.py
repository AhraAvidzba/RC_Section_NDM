from pathlib import Path
import importlib.util
import difflib

directory=Path(__file__).resolve().parent
root=directory.parents[3]
spec=importlib.util.spec_from_file_location('reader',root/'tools/build_workbook/Verify-SectionRecipeGuide.py')
reader=importlib.util.module_from_spec(spec);spec.loader.exec_module(reader)
old,new=reader.read(directory/'Baseline.xlsm'),reader.read(directory/'Candidate.xlsm')
for key,value in new['names'].items():
    if old['names'].get(key)!=value:
        print('NAME',key,old['names'].get(key),value)
for address,cell in old['sheets']['Справка']['cells'].items():
    if 'Rebar.Loc2row' in (cell['value'] or '') or 'Rebar.Loc3row' in (cell['value'] or ''):
        print('HELP',address,cell['value'])
for sheet in ('Справка','Config'):
    links=old['sheets'][sheet]['document'].find(reader.S+'hyperlinks')
    if links is not None:
        print('LINKS',sheet,[e.attrib for e in links if 'Config' in (e.get('location') or '')][:8])
name='modTestShapeArchitecture.DumpShapeArchitectureSnapshot.txt'
a=(directory/'BaselineGeometry'/name).read_text(encoding='utf-8-sig').splitlines()
b=(directory/'CandidateGeometry'/name).read_text(encoding='utf-8-sig').splitlines()
changes=[(x,y) for x,y in zip(a,b) if x!=y]
print('CHANGED_LINES',len(changes),'LINE_COUNTS',len(a),len(b))
for x,y in changes[:6]:
    print('BEFORE',x[:500]);print('AFTER',y[:500])
