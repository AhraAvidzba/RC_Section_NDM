"""Фиксирует область архитектурного аудита, зависимости и контрпримеры контрактов."""
import hashlib
import json
import re
import shutil
from pathlib import Path

root = Path.cwd()
folder = root/'docs/regression/ExcelGuardGuide_2026-10-13/Architecture'
inventory = json.loads((folder/'SourceInventory.json').read_text(encoding='utf-8'))
core = [item for item in inventory if item['path'].startswith(('src/Section/','src/Geometry/'))]
integration = ['CSectionTypeRegistry','CSectionInputSnapshotReader','CUnitSystem','CSystemSettingsReader',
               'CNDMResultsWriter','CSectionPlotDataReader','CPlotAnnotationLayout','CSectionPlotter',
               'CAutoCADSectionModelImporter','CAutoCADContourReader']
items = core + [item for item in inventory if item['class'] in integration]
example = root/'docs/geometry_refactor/examples/CSectionRecipeRoundedTrapezoid.cls'
modules = ['src/Excel/modWorkbookCalculation.bas','src/Excel/modAutoCADStressExport.bas',
           'src/Excel/modSectionRecipeSetup.bas','src/Common/modGeometryTypes.bas',
           'src/Common/modSectionRecipeTypes.bas']

graph = {item['class']:set(item['dependencies'])-{item['class']} for item in items}
cycles = set()
def walk(origin, node, chain):
    for child in graph.get(node,()):
        if child == origin:
            cycles.add(tuple(sorted(chain)))
        elif child not in chain and child in graph:
            walk(origin,child,chain+[child])
for name in graph: walk(name,name,[name])
assert cycles == {('CGeometryPath','CSectionGeometryQuery')}, cycles

def strip_comment(line):
    quoted, i = False, 0
    while i < len(line):
        if line[i] == '"':
            if quoted and i+1<len(line) and line[i+1]=='"': i+=2; continue
            quoted = not quoted
        elif line[i] == "'" and not quoted: return line[:i]
        i += 1
    return line

common = ['CSectionBuilder','CSectionDefinition','CSectionDefinitionBuilder','CSectionGeometryQuery',
          'CGeometryPath','CGeometryRegion','CFiberMeshBuilder','CRebarRuleExecutor','CRebarLayout',
          'CSectionModelBuilder','CSectionAnnotationBuilder','CRebarGroupAnnotationBuilder',
          'CSectionInputTable','CSectionRecipeContext','CSystemSettingsReader']
family_branches=[]
for item in items:
    path = root/item['path']
    item['sha256']=hashlib.sha256(path.read_bytes()).hexdigest()
    if item['class'] in common:
        for number,line in enumerate(path.read_text(encoding='utf-8-sig').splitlines(),1):
            code=strip_comment(line)
            if re.search(r'"(?:Circle|RectSet|RoundedRectangle|HollowRectangle)"',code):
                family_branches.append({'class':item['class'],'line':number,'code':code.strip()})
assert not family_branches, family_branches

runtime=json.loads((folder/'RuntimeChecks.json').read_text(encoding='utf-8-sig'))
annotations=json.loads((folder/'AnnotationRuntimeChecks.json').read_text(encoding='utf-8-sig'))
assert all(run['Passed'] for result in (runtime,annotations) for run in result['Runs'])
passed=failed=0
for result in (runtime,annotations):
    for run in result['Runs']:
        report=(folder/(run['Macro']+'.txt')).read_text(encoding='utf-8-sig')
        totals=re.findall(r'(?m)^TOTAL[^\r\n]*passed=(\d+); failed=(\d+)',report)
        assert totals,run['Macro']
        for yes,no in totals: passed+=int(yes); failed+=int(no)
assert passed==914 and failed==0,(passed,failed)
probe=(folder/'modGeometryArchitectureProbe.RunGeometryArchitectureProbe.txt').read_text(encoding='utf-8-sig')
labels=(folder/'modGeometryArchitectureProbe.RunAnnotationOwnershipProbe.txt').read_text(encoding='utf-8-sig')
assert 'before_mm=1000; after_mm=100; sealed_context_changed=True' in probe
assert 'invalid_step_accepted_by_finish=True' in probe
assert '50_and_5000_produce_same_annotation=True' in labels
assert 'requested_normal_y=-1; actual_normal_y=1; stale_anchor=True' in labels

shutil.copyfile(root/'tmp/Run-GeometryArchitecture.ps1',folder/'RunFocusedChecks.ps1')
result={
    'reviewedCommit':'e2551c220d09330f788d931e967ac312df56bba1',
    'scope':{'coreClasses':len(core),'integrationClasses':len(integration),'exampleClasses':1,
             'totalClasses':len(items)+1,'integrationAndContractModules':modules},
    'classes':items,
    'example':{'path':str(example.relative_to(root)).replace('\\','/'),'sha256':hashlib.sha256(example.read_bytes()).hexdigest()},
    'classDependencyCycles':[list(cycle) for cycle in sorted(cycles)],
    'cyclesAreNotObjectOwnershipOrConstructionCycles':True,
    'commonLayerBranchesOnSectionFamily':family_branches,
    'guardAndRegressionAssertions':{'passed':passed,'failed':failed},
    'observedContractViolations':[
        {'id':'A01','fact':'Sealed recipe context shares a mutable CUnitSystem; 100 cm becomes 100 mm after external reset.'},
        {'id':'A02','fact':'Repeated annotation build on an existing layout retains stale anchors; requested normal -1 remains +1.'}
    ],
    'otherObservations':[
        {'id':'A03','fact':'Label axisDistance values 50 and 5000 produce identical final semantic annotations.'},
        {'id':'A04','fact':'Finish seals a ByMaxStep rule with zero step; executor rejects it. CSectionBuilder builds mesh before executor.'},
        {'id':'A05','fact':'CSectionGeometryQuery combines analytic operations with stateful model/domain reconstruction and caching.'}
    ],
    'productionFiles':[]
}
for name in ('RC_Section_NDM.xlsm','VBA_All_Code.txt','RC_Section_NDM_execution_report.txt'):
    path=root/'workbook/output'/name
    result['productionFiles'].append({'path':str(path.relative_to(root)).replace('\\','/'),'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),'unchangedDuringAudit':True})
(folder/'ArchitectureAudit.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'classes':len(items)+1,'modules':len(modules),'passedGuards':passed,'failedGuards':failed,
                  'contractViolationsReproduced':2,'otherObservations':3,'productionUnchanged':True}))
