"""Собирает проверенные доказательства без тестовых XLSM и пользовательских чертежей."""
import hashlib
import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

directory = Path(__file__).resolve().parent
root = directory.parents[3]

def read(name):
    return json.loads((directory/name).read_text(encoding='utf-8-sig'))

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

checks = {
    'arbitrary_geometry': 'CandidateFinalDirected/Run.json',
    'legacy_and_settings': 'CandidateDirected01/Run.json',
    'native_rotation_bounds_repeat': 'NativeAngleAcceptance.json',
    'legacy_geometry': 'LegacyGeometryEquality.json',
    'legacy_pixels': 'LegacyPreviewEquality.json',
    'ordinary_AutoCAD_2023': 'DocumentScopedCAD02/Run.json',
    'tapered_facet_screenshot': 'TaperedFacetAcceptance.json',
    'published_source_equality': 'PublishedSourceEquality.json',
    'published_data_preservation': 'PublishedPreservation.json',
    'standalone_help': 'PublishedStandaloneHelp.json',
}
for name, path in checks.items():
    result = read(path)
    assert result.get('Passed', result.get('status') == 'STANDALONE_HELP_PASS'), (name,path)
publication = read('Publication.json')
assert not publication['ReadbackPending']
assert publication['PublishedCodeMatchesCandidate'] and publication['AllCodeMatchesActualWorkbookExport']
assert sha(root/'workbook/output/RC_Section_NDM.xlsm').upper() == publication['WorkbookSHA256']
assert sha(root/'workbook/output/VBA_All_Code.txt').upper() == publication['AllCodeSHA256']
assert sha(root/'workbook/output/RC_Section_NDM_execution_report.txt').upper() == read('Before.json')['ExecutionReportSHA256']

groups = {
    'NativeExports.zip': [directory/name for name in ('BaselineNativeVBA.txt','PublishedNativeVBA.txt')],
    'NativeAngles.zip': [path for stage in ('BaselineNativeAngles','CandidateNativeAngles')
                         for path in (directory/stage).iterdir() if path.suffix.lower() in ('.png','.xlsx','.tsv','.json','.txt')],
    'LegacyGeometryAndPreviews.zip': [path for stage in ('BaselineGeometry','CandidateGeometry','BaselineLegacyPreviews','CandidateLegacyPreviews')
                                      for path in (directory/stage).iterdir() if path.suffix.lower() in ('.png','.tsv','.json','.txt')],
    'DirectedAndCADLogs.zip': [path for stage in ('CandidateDirected01','CandidateFinalDirected','DocumentScopedCAD02')
                               for path in (directory/stage).iterdir() if path.suffix.lower() in ('.json','.txt','.dwg')],
    'TaperedFacetEvidence.zip': [path for path in (directory/'TaperedFacet02').iterdir()
                                if path.suffix.lower() in ('.tsv','.json')],
}
archives = []
for name, paths in groups.items():
    target = directory/name
    with ZipFile(target,'w',ZIP_DEFLATED,compresslevel=9) as archive:
        for path in sorted(paths):
            assert path.suffix.lower() != '.xlsm'
            archive.write(path,path.relative_to(directory).as_posix())
    archives.append({'File':name,'Bytes':target.stat().st_size,'SHA256':sha(target),'Files':len(paths)})

acceptance = {
    'Passed': True,
    'BaselineHead': read('Before.json')['Head'],
    'AnglesDegrees': [0,30,45,90,135,225,315],
    'Requirements': {
        'ExactProjectionSupportArbitraryNormal': checks['arbitrary_geometry'],
        'CircleRoundingHollowRectangleOpeningBoundary': checks['arbitrary_geometry'],
        'ReadableRotationInActualChartCoordinates': checks['native_rotation_bounds_repeat'],
        'TextBoundsClearanceSeparationRepeat': checks['native_rotation_bounds_repeat'],
        'CardinalPlacementAndTypographyExact': checks['native_rotation_bounds_repeat'],
        'LegacySettingsUnitsGroupingAndSnapshot': checks['legacy_and_settings'],
        'PhysicalModelsAndExistingAnnotationsUnchanged': checks['legacy_geometry'],
        'NineExistingNativeSchemasPixelIdentical': checks['legacy_pixels'],
        'AvailableAutoCADGeometryRegression': checks['ordinary_AutoCAD_2023'],
        'RequestedFacetImageWithExistingAnnotations': checks['tapered_facet_screenshot'],
        'WorkbookSourceAndAllCodeSynchronized': checks['published_source_equality'],
        'UserDataAndSavedSchemaPreserved': checks['published_data_preservation'],
        'UpdatedVbaGuideAndHelpStandalone': checks['standalone_help'],
    },
    'Publication': publication,
    'UserExecutionReportUnchanged': True,
    'FullOnOffRun': False,
    'NewClassesOrParallelAPI': False,
    'TemporaryHelpersInProduction': False,
    'AutoCADAngleTextTested': False,
    'Report': 'docs/reports/AngleAnnotations_2026-10-14.md',
    'Screenshot': 'TaperedFacet02/RoundedTapered.png',
    'Archives': archives,
}
(directory/'Acceptance.json').write_text(json.dumps(acceptance,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'Passed':True,'Requirements':len(acceptance['Requirements']),'Archives':archives},ensure_ascii=False))
