"""Проверяет весь сохранённый VBA, зависимости удалённых механизмов и границы правки."""
import hashlib
import importlib.util
import json
import re
import subprocess
from pathlib import Path
from zipfile import ZipFile

directory = Path(__file__).resolve().parent
root = directory.parents[3]
spec = importlib.util.spec_from_file_location('presentation', root/'tools/build_workbook/Verify-GuidePresentation.py')
presentation = importlib.util.module_from_spec(spec)
spec.loader.exec_module(presentation)
native = presentation.components(directory/'FinalNativeVBA.txt')
baseline = json.loads((directory/'BaselineAudit.json').read_text(encoding='utf-8'))
retired = [name for name in baseline['BeforeRemovalDependencies'] if name != 'HeaderRow'] + [
    'Solver.DirectState.DiagramExtension', 'AutoCAD.Export.PrincipalAxesEnabled', 'mMigrationWarning',
]
errors = []
for name, component in native.items():
    for symbol in retired:
        if re.search(r'\b'+re.escape(symbol)+r'\b', component[2], re.I):
            errors.append('Retired symbol in '+name+': '+symbol)
assert len(native) == 157

# Преобразование допустимых значений и чтение поддерживаемых форматов остаются.
assert 'GetRequiredBoolean("General.DiagramExtension")' in native['CMaterialModelProvider'][2]
assert 'ValidateDiagramExtensionSetting True' in native['CSystemSettingsReader'][2]
assert 'SetLegacyContourTag' in native['CFakeAcadContour'][2]
assert 'CheckMixedContourSources' in native['modTestAutoCADContours'][2]
setup = native['modSectionRecipeSetup'][2]
tables = set(re.findall(r'RequiredTable\("([^"]+)"\)', setup))
assert tables == {'rngSystemSettings', 'rngSectionRecipeTypes'}
assert not re.search(r'\.Insert\b|\.Worksheets\.Add\b|\.Names\.Add\b', setup, re.I)
assert 'UpdateSectionRecipeTypes' in setup
# HeaderRow в удалённом updater-е был частным помощником миграции.
# Одноимённые действующие свойства входных таблиц и отчётов удалять нельзя.
assert not re.search(r'\bHeaderRow\b', setup, re.I)

with ZipFile(directory/'Candidate.xlsm') as archive:
    for name in archive.namelist():
        if name.endswith(('.xml', '.vml')):
            text = archive.read(name).decode('utf-8', errors='replace')
            if 'UpdateSectionRecipeConfig' in text:
                errors.append('Retired macro instruction/action in '+name)

# Все производственные изменения ограничены Excel-адаптерами, гайдом и комментарием provider-а.
allowed = {
    'src/Excel/CExecutionReport.cls', 'src/Excel/CPlotAnnotationLayout.cls',
    'src/Excel/CSectionPlotDataReader.cls', 'src/Excel/CSectionPlotter.cls',
    'src/Excel/CSystemSettingsReader.cls', 'src/Excel/modAutoCADStressExport.bas',
    'src/Excel/modSectionRecipeSetup.bas', 'src/Materials/CMaterialModelProvider.cls',
    'src/Common/modSectionRecipeApiGuide.bas',
}
source_inventory = []
for path in sorted((root/'src').rglob('*')):
    if path.suffix not in ('.bas', '.cls'): continue
    relative = path.relative_to(root).as_posix()
    old = subprocess.check_output(['git', 'show', baseline_ref := '9dbbba21536d886a71db878469b448886a1853a2:'+relative], cwd=root).decode('utf-8-sig').replace('\r\n','\n')
    current = path.read_text(encoding='utf-8-sig')
    if old != current and relative not in allowed: errors.append('Unexpected production change: '+relative)
    source_inventory.append({'Path': relative, 'Unchanged': old == current, 'SHA256': hashlib.sha256(path.read_bytes()).hexdigest()})

def executable(text):
    # Убирает только комментарии за пределами строковых литералов и пустые строки.
    parts = re.findall(r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27\n]+|\n', text)
    return '\n'.join(line.strip() for line in ''.join(p for p in parts if not p.startswith("'")).splitlines() if line.strip())

provider = 'src/Materials/CMaterialModelProvider.cls'
old_provider = subprocess.check_output(['git','show','9dbbba21536d886a71db878469b448886a1853a2:'+provider], cwd=root).decode('utf-8-sig')
assert executable(old_provider) == executable((root/provider).read_text(encoding='utf-8-sig'))
assert len(source_inventory) == 110
assert all(item['Unchanged'] for item in source_inventory if any('/'+group+'/' in item['Path'] for group in ('Geometry','Section','Solver','Crack','Stability','Batch','Units')))
result = {
    'Passed': not errors, 'NativeComponentsAudited': len(native),
    'BaselineProceduresInventoried': sum(len(c['Procedures']) for c in baseline['Components']),
    'BaselineStructuralWriteAndDynamicCallCandidates': sum(len(c['StructureOrWriteCandidates']) for c in baseline['Components']),
    'RetiredSymbolsAbsent': retired, 'ScopedRetiredSymbolsAbsent': {'modSectionRecipeSetup': ['HeaderRow']},
    'TypeUpdaterTables': sorted(tables),
    'SupportedReadersImportExportAndDWGChecksRetained': True,
    'CalculationAndGeometryAlgorithmsUnchanged': True,
    'SourceInventory': source_inventory, 'Errors': errors,
}
(directory/'FinalAudit.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps({k:v for k,v in result.items() if k != 'SourceInventory'},ensure_ascii=False))
if errors: raise SystemExit(1)
