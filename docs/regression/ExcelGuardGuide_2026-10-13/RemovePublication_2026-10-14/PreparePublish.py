"""Готовит штатную публикацию с проверками именно этого изменения API."""
from pathlib import Path

directory = Path(__file__).resolve().parent
previous = directory.parent/'AngleAnnotations_2026-10-14'
old_path = 'docs/regression/ExcelGuardGuide_2026-10-13/AngleAnnotations_2026-10-14'
new_path = 'docs/regression/ExcelGuardGuide_2026-10-13/RemovePublication_2026-10-14'
source = (previous/'Publish.ps1').read_text(encoding='utf-8-sig').replace(old_path, new_path)
start = source.index("$native=Receipt 'NativeAngleAcceptance.json'")
end = source.index("$candidate=Join-Path", start)
source = source[:start] + """$equivalence=Receipt 'Equivalence.json'
$directed=Receipt 'Directed01/Run.json'
$guide=Receipt 'GuideNativeAccepted/GuideEvidence.json'
if(-not $preservation.Passed -or -not $preservation.NativeChangesRestricted -or -not $preservation.NoEmbeddedMarkers -or -not $preservation.UserDataFormulasStylesNamesAndDrawingsPreserved -or -not $equality.Passed -or $help.status -ne 'STANDALONE_HELP_PASS' -or -not $equivalence.Passed -or $equivalence.Calls -ne 13 -or -not $directed.Passed -or $guide.Status -ne 'DIRECTED_NATIVE_GUIDE_PASS'){throw 'Publication options acceptance is incomplete.'}
""" + source[end:]
(directory/'Publish.ps1').write_text(source, encoding='utf-8-sig')
source = (previous/'Finalize.ps1').read_text(encoding='utf-8-sig').replace(old_path,new_path)
(directory/'Finalize.ps1').write_text(source, encoding='utf-8-sig')
print('PUBLICATION_SCRIPTS_READY')
