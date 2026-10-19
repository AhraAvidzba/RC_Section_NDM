"""Строгая адресная приёмка упрощения рецептов: неизменные снимки и состав проверок."""
import argparse
import hashlib
import json
import re
import subprocess
from io import BytesIO
from collections import Counter
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
D = ROOT / 'docs/regression/ExcelGuardGuide_2026-10-13/RecipeSimplification_2026-10-16'


def data(name):
    path = D / name
    if path.exists():
        return path.read_bytes()
    with ZipFile(D / 'Evidence.zip') as archive:
        return archive.read(name)


def text(name):
    return data(name).decode('utf-8-sig')


def receipt(name):
    return json.loads(text(name))


def verify(published=False):
    comparisons, runs = [], {}

    def compare(before, after):
        a, b = data(before), data(after)
        assert a == b, 'Snapshot differs: '+before+' / '+after
        comparisons.append({'Before': before, 'After': after, 'Bytes': len(a),
                            'SHA256': hashlib.sha256(a).hexdigest()})

    def run(stage, macros):
        record = receipt(stage+'/Run.json')
        assert record['Passed'] and record['ProductionUnchanged'], stage
        assert [item['Macro'] for item in record['Runs']] == macros, stage
        assert all(item['Passed'] for item in record['Runs']), stage
        for item in record['Runs']:
            body = text(stage+'/'+item['Macro']+'.txt')
            assert len(body.strip()) > 0
            assert not re.search(r'(?m)^FAIL[: ]|failed=[1-9]|RUNTIME ERROR:|^ERROR:', body), stage
            if not item['Snapshot']:
                assert re.search(r'(?m)^TOTAL[^\r\n]*failed=0\b', body), stage
        runs[stage] = record

    matrices = ['modTestSectionRecipes.Dump'+name+'Recipe'+suffix for name, suffix in
                [('Circle', 'Matrix'), ('RectSet', 'Matrix'), ('Rounded', 'Matrix'),
                 ('Hollow', 'Matrix'), ('Hollow', 'EdgeMatrix')]]
    contract = 'modTestAudit02Api.DumpAudit02CloseoutInputContract'
    rules = 'modTestAudit02Snapshots.DumpRecipeRuleContracts'
    calculations = 'modTestAudit02Snapshots.DumpAudit02Calculations'
    plots = ['modTestSectionRecipes.Dump'+name+'RecipePreview' for name in ('Circle', 'RectSet', 'Rounded', 'Hollow')]
    states = ['modTestSectionRecipes.Dump'+name+'RecipeCalculations' for name in ('Circle', 'RectSet', 'Rounded', 'Hollow')]
    directed = ['modTestSectionRecipes.'+name for name in ('RunWorkbookRecipeInputTests', 'RunRecipeDefinitionTests',
                'RunCircleRecipeInputTests', 'RunRoundedRecipeInputTests', 'RunHollowRecipeInputTests',
                'RunCircleBindingTests', 'RunCircleBindingWorkbookTests', 'RunRecipeAnnotationCountTests', 'RunRecipeRuleTests')]
    directed += ['modTestRecipeSetbacks.RunRecipeSetbackUnitTests', 'modTestRecipeSetbacks.RunRecipeSetbackInputTests',
                 'modTestRecipeOperations.RunRecipeOperationTests', 'modTestRecipeOperations.RunRecipeDiagnosticTests',
                 'modTestAudit02Api.RunAudit02CloseoutRectSetInputTests', 'modTestAudit02Api.RunAudit02CloseoutContactTests',
                 'modTestAudit02Api.RunRecipeContactOptionsContract', 'modTestRecipeApiCatalogue.RunAdvancedGuideExamples']
    run('BaselineMatricesNative', matrices+[contract])
    run('CandidateMatrices', matrices+[contract, rules])
    run('AcceptedMatrices', matrices+[contract, rules])
    assert runs['AcceptedMatrices']['SourceSHA256'].lower() == receipt('Package.json')['AcceptedSHA256']
    run('BaselineContractsVerified', [rules, calculations])
    run('BaselineSP63', [calculations])
    run('CandidateSP35', [calculations])
    run('CandidateSP63', [calculations])
    run('BaselinePlotRun', plots)
    run('CandidatePlotRun', plots)
    run('CandidateDirected', directed)
    run('AcceptedSmoke', [directed[0], 'modTestSectionRecipes.RunCircleBindingWorkbookTests', 'modTestAudit02Api.RunRecipeContactOptionsContract'])
    for norm in ('35', '63'):
        run('BaselineStates'+norm, states)
        run('CandidateStates'+norm, states)
        for macro in states:
            compare('BaselineStates'+norm+'/'+macro+'.txt', 'CandidateStates'+norm+'/'+macro+'.txt')
    for macro in matrices+[contract]:
        compare('BaselineMatricesNative/'+macro+'.txt', 'CandidateMatrices/'+macro+'.txt')
        compare('BaselineMatricesNative/'+macro+'.txt', 'AcceptedMatrices/'+macro+'.txt')
    compare('BaselineContractsVerified/'+rules+'.txt', 'CandidateMatrices/'+rules+'.txt')
    compare('BaselineContractsVerified/'+rules+'.txt', 'AcceptedMatrices/'+rules+'.txt')
    compare('BaselineContractsVerified/'+calculations+'.txt', 'CandidateSP35/'+calculations+'.txt')
    compare('BaselineSP63/'+calculations+'.txt', 'CandidateSP63/'+calculations+'.txt')
    for macro in plots:
        compare('BaselinePlotRun/'+macro+'.txt', 'CandidatePlotRun/'+macro+'.txt')
    fixtures = ['Circle', 'Rectangle', 'LSection', 'TwoRectangles', 'RoundedSimple', 'RoundedTapered',
                'RoundedMixed', 'RoundedMixedLeft', 'HollowCentered', 'HollowOffset']
    for fixture in fixtures:
        compare('BaselinePlots/'+fixture+'.png', 'CandidatePlots/'+fixture+'.png')
    state_rows = [line.split('\t') for norm in ('35', '63') for macro in states
                  for line in text('CandidateStates'+norm+'/'+macro+'.txt').splitlines()]
    nds = [row for row in state_rows if row[1] == 'NDS']
    assert len(nds) == 240 and all(row[3] == 'b:1' for row in nds)
    assert len([row for row in state_rows if row[1] == 'BATCH']) == 40
    assert len(re.findall(r'(?m)^CASE\|', text('CandidateMatrices/'+rules+'.txt'))) == 60
    rule_type = re.search(r'Public Type TRebarRule\n(.*?)End Type', (ROOT/'src/Common/modSectionRecipeTypes.bas').read_text(encoding='utf-8-sig'), re.S).group(1)
    fields = re.findall(r'^\s*(\w+) As ', rule_type, re.M)
    serializer = (ROOT/'tests/modTestAudit02Snapshots.bas').read_text(encoding='utf-8-sig').split('Private Function RecipeRuleFields', 1)[1]
    assert all(re.search(r'\brule\.'+field+r'\b', serializer) for field in fields)
    assert receipt('AcceptedGuidePresentation.json')['Passed'] and receipt('AcceptedSourceEquality.json')['Passed']
    assert receipt('Package.json')['Status'] == 'RECIPE_PACKAGE_PASS'
    assert receipt('CommentAudit.json')['status'] == 'PASS'
    # Старые Public сигнатуры сохраняются; общий помощник добавляется отдельным членом.
    before_api = json.loads(subprocess.check_output(['git', 'show', '2e397426:docs/geometry_refactor/SECTION_RECIPE_API_MANIFEST.json'], cwd=ROOT))
    after_api = json.loads((ROOT/'docs/geometry_refactor/SECTION_RECIPE_API_MANIFEST.json').read_text(encoding='utf-8-sig'))
    a, b = [{member['id']: member for member in item['members']} for item in (before_api, after_api)]
    assert b.keys() - a.keys() == {'modSectionRecipeRows.RecipeContactOptions'} and not a.keys() - b.keys()
    assert all({k: v for k, v in a[name].items() if k != 'line'} == {k: v for k, v in b[name].items() if k != 'line'} for name in a)
    with ZipFile(BytesIO(data('BaselineSource.zip'))) as archive:
        old_runtime = {name: archive.read(name).decode('utf-8-sig').replace('\r\n', '\n')
                       for name in archive.namelist() if name.startswith('src/') and name.endswith(('.bas', '.cls'))}
        current_runtime = {path.relative_to(ROOT).as_posix(): path.read_text(encoding='utf-8-sig')
                           for path in (ROOT/'src').rglob('*') if path.suffix in ('.bas', '.cls')}
        assert current_runtime.keys() - old_runtime.keys() == {'src/Geometry/modSectionRecipeRows.bas'}
        assert not old_runtime.keys() - current_runtime.keys()
        changed_runtime = {name for name in old_runtime if old_runtime[name] != current_runtime[name]}
        assert changed_runtime == {'src/Geometry/CSectionRecipe'+name+'.cls' for name in ('Circle', 'RectSet', 'PointedRect', 'HollowRectangle')} | {'src/Common/modSectionRecipeApiGuide.bas'}
        def executable(source):
            return ''.join(part if part.startswith('"') else re.sub(r'\s+', '', part.lower())
                           for part in re.findall(r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27]+', source)
                           if not part.startswith("'"))
        new_methods = {'tests/modTestAudit02Api.bas': ['RunRecipeContactOptionsContract'],
                       'tests/modTestAudit02Snapshots.bas': ['DumpRecipeRuleContracts', 'ConfigureCircleRuleContract', 'RecipeRuleFields']}
        for name in (name for name in archive.namelist() if name.startswith('tests/') and name.endswith(('.bas', '.cls'))):
            current = (ROOT/name).read_text(encoding='utf-8-sig')
            for method in new_methods.get(name, []):
                current, removed = re.subn(r'(?ms)^(?:Public|Private) (Function|Sub) '+method+r'\(.*?^End \1\b', '', current)
                assert removed == 1
            assert executable(archive.read(name).decode('utf-8-sig')) == executable(current), 'Existing tests changed: '+name
    passed = sum(int(value) for macro in directed for value in re.findall(r'(?m)^TOTAL[^\r\n]*passed=(\d+)', text('CandidateDirected/'+macro+'.txt')))
    if published:
        publication = receipt('Publication.json')
        assert publication['Status'] == 'PUBLISHED_VERIFIED' and publication['AllCodeMatchesPublishedNative']
        stage = 'GuideReadability/' if publication.get('FinalGuideReadability') else ''
        assert receipt(stage+'PublishedSourceEquality.json')['Passed'] and receipt(stage+'PublishedGuidePresentation.json')['Passed']
        accepted_sha = receipt('Package.json')['AcceptedSHA256']
        if stage:
            final_package = receipt(stage+'Package.json')
            assert final_package['AllExecutableVBAUnchanged'] and final_package['AllNonVBAPartsByteIdentical']
            assert final_package['BeforeWorkbookSHA256'] == accepted_sha
            accepted_sha = final_package['AcceptedSHA256']
        assert hashlib.sha256((ROOT/'workbook/output/RC_Section_NDM.xlsm').read_bytes()).hexdigest() == publication['WorkbookSHA256'].lower() == accepted_sha
        assert (ROOT/'workbook/output/VBA_All_Code.txt').read_bytes() == data(stage+'PublishedNativeVBA.txt')
        assert hashlib.sha256((ROOT/'workbook/output/RC_Section_NDM_execution_report.txt').read_bytes()).hexdigest() == receipt('Before.json')['UserReportSHA256'].lower()
    return {'Status': 'RECIPE_SIMPLIFICATION_PASS', 'DirectedSuites': len(directed), 'DirectedAssertions': passed,
            'ExactComparisons': comparisons, 'RuleDefinitionCases': 60, 'RuleFieldsCovered': len(fields),
            'NDSStatesConverged': len(nds), 'ElementStates': dict(Counter(row[1] for row in state_rows)),
            'BatchCombinations': 40, 'PlotsIdentical': len(fixtures), 'ExistingPublicAPIUnchanged': True,
            'ExistingRuntimeSourcesCompared': len(old_runtime), 'ChangedExistingRuntimeSources': sorted(changed_runtime),
            'AddedRuntimeModules': ['modSectionRecipeRows'], 'ExistingTestsUnchanged': True,
            'RunStages': list(runs), 'FullOnOffRun': False, 'TolerancesChanged': False, 'PublishedOutputVerified': published}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true')
    parser.add_argument('--published', action='store_true')
    args = parser.parse_args()
    result = verify(args.published)
    (D/'Acceptance.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps({key: value for key, value in result.items() if key != 'ExactComparisons'}, ensure_ascii=False))
