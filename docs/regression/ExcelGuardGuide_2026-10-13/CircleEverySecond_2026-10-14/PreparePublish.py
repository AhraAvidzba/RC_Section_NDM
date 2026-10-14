from pathlib import Path
directory=Path(__file__).resolve().parent;previous=directory.parent/'RemovePublication_2026-10-14'
old='docs/regression/ExcelGuardGuide_2026-10-13/RemovePublication_2026-10-14'
new='docs/regression/ExcelGuardGuide_2026-10-13/CircleEverySecond_2026-10-14'
source=(previous/'Publish.ps1').read_text(encoding='utf-8-sig').replace(old,new)
start=source.index("$equivalence=Receipt 'Equivalence.json'")
end=source.index('$candidate=Join-Path',start)
source=source[:start]+"""$directed=Receipt 'Directed01/Run.json'
$finalDirected=Receipt 'FinalWorkbookDirected/Run.json'
$factory=Receipt 'FactoryTest.json'
$migration=Receipt 'RecipeConfigUpdate.json'
if(-not $preservation.Passed -or -not $preservation.NativeChangesRestricted -or -not $preservation.NoEmbeddedMarkers -or -not $preservation.UserDataFormulasStylesNamesAndDrawingsPreserved -or -not $preservation.ModelsEqualExcept72NewInputCells -or -not $equality.Passed -or $help.status -ne 'STANDALONE_HELP_PASS' -or -not $directed.Passed -or -not $finalDirected.Passed -or -not $factory.Passed -or -not $migration.Idempotent){throw 'Circle binding acceptance is incomplete.'}
"""+source[end:]
(directory/'Publish.ps1').write_text(source,encoding='utf-8-sig')
(directory/'Finalize.ps1').write_text((previous/'Finalize.ps1').read_text(encoding='utf-8-sig').replace(old,new),encoding='utf-8-sig')
print('PUBLICATION_SCRIPTS_READY')
