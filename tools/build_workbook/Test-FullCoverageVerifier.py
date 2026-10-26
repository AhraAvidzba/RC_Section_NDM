"""Проверяет, что новая обязательная приёмка не теряется при сводном PASS."""
import copy
import importlib.util
import json
from pathlib import Path
import re
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

spec = importlib.util.spec_from_file_location('coverage', Path(__file__).with_name('Summarize-FullOnOffAcceptance.py'))
coverage = importlib.util.module_from_spec(spec)
spec.loader.exec_module(coverage)


class TestFullCoverage(unittest.TestCase):
    """Исторический состав сохраняется; версия 6 требует новую и прежнюю геометрию."""

    def manifest_for_version(self, version):
        """Создаёт только синтетический план для проверки состава, без Excel и расчётов."""
        counts = {
            1: coverage.GROUP_COUNTS,
            2: coverage.CURRENT_GROUP_COUNTS,
            3: coverage.INDIVIDUAL_ROW_GROUP_COUNTS,
            4: coverage.SNAPSHOT_ACCELERATION_GROUP_COUNTS,
            5: coverage.MESSAGE_PRESENTATION_GROUP_COUNTS,
            6: coverage.CONTOUR_ELEMENT_GROUP_COUNTS,
        }[version]
        feature_macros = {
            4: coverage.SNAPSHOT_ACCELERATION_MACROS,
            5: coverage.MESSAGE_PRESENTATION_MACROS,
            6: coverage.CONTOUR_ELEMENT_MACROS,
        }.get(version, set())
        manifest = dict(CoverageDefinition='AllActiveAcceptanceIncludingFocused',
                        CoverageVersion=version, IncludesFocusedAcceptance=True, Jobs=[])
        for group, count in counts.items():
            for mode in ('No', 'Yes'):
                if group == 'Full':
                    macros = []
                elif group == 'CurrentFeatures':
                    macros = sorted(feature_macros)
                else:
                    macros = sorted(coverage.INDIVIDUAL_ROW_MACROS) if group == 'Geometry' and version >= 3 else []
                    macros += ['modTestCoverage.TestSyntheticCase' + str(i) for i in range(count - len(macros))]
                manifest['Jobs'].append(dict(Group=group, Mode=mode, ExpectedSuites=count, Macro=macros))
        return manifest

    def setUp(self):
        self.historical = dict(CoverageDefinition='AllActiveAcceptanceIncludingFocused',
                               CoverageVersion=3, IncludesFocusedAcceptance=True, Jobs=[])
        for group, count in coverage.INDIVIDUAL_ROW_GROUP_COUNTS.items():
            for mode in ('No', 'Yes'):
                macros = []
                if group == 'Geometry':
                    macros = sorted(coverage.INDIVIDUAL_ROW_MACROS)
                    macros += ['modTestCoverage.TestSyntheticCase' + str(i) for i in range(count - len(macros))]
                self.historical['Jobs'].append(dict(Group=group, Mode=mode, ExpectedSuites=count, Macro=macros))
        self.current = copy.deepcopy(self.historical)
        self.current['CoverageVersion'] = 4
        for mode in ('No', 'Yes'):
            self.current['Jobs'].append(dict(Group='CurrentFeatures', Mode=mode, ExpectedSuites=3,
                                            Macro=sorted(coverage.SNAPSHOT_ACCELERATION_MACROS)))

    def test_historical_definition_remains_supported(self):
        self.assertEqual(coverage.coverage_errors(self.historical), [])

    def test_all_historical_versions_keep_their_original_counts(self):
        for version, expected in ((1, 132), (2, 144), (3, 156), (4, 162), (5, 168)):
            with self.subTest(version=version):
                manifest = self.manifest_for_version(version)
                self.assertEqual(sum(job['ExpectedSuites'] for job in manifest['Jobs']), expected)
                self.assertEqual(coverage.coverage_errors(manifest), [])

    def test_v6_has_178_suites_including_eleven_features_per_mode(self):
        manifest = self.manifest_for_version(6)
        self.assertEqual(coverage.coverage_errors(manifest), [])
        self.assertEqual(sum(job['ExpectedSuites'] for job in manifest['Jobs']), 178)
        features = [job for job in manifest['Jobs'] if job['Group'] == 'CurrentFeatures']
        self.assertEqual(len(features), 2)
        self.assertEqual([job['ExpectedSuites'] for job in features], [11, 11])

    def test_each_new_v6_macro_is_mandatory_in_each_mode(self):
        added = coverage.CONTOUR_ELEMENT_MACROS - coverage.MESSAGE_PRESENTATION_MACROS
        self.assertEqual(len(added), 5)
        for macro in added:
            for mode in ('No', 'Yes'):
                with self.subTest(macro=macro, mode=mode):
                    broken = self.manifest_for_version(6)
                    job = next(job for job in broken['Jobs'] if job['Group'] == 'CurrentFeatures' and job['Mode'] == mode)
                    job['Macro'].remove(macro)
                    job['ExpectedSuites'] -= 1
                    self.assertTrue(coverage.coverage_errors(broken))

    def test_v6_macro_substitution_duplicate_and_extra_are_rejected(self):
        for mutation in ('substitute', 'duplicate', 'extra'):
            with self.subTest(mutation=mutation):
                broken = self.manifest_for_version(6)
                job = broken['Jobs'][-1]
                if mutation == 'substitute':
                    job['Macro'][0] = 'modTestCoverage.TestUnexpectedCase'
                elif mutation == 'duplicate':
                    job['Macro'][0] = job['Macro'][1]
                else:
                    job['Macro'].append('modTestCoverage.TestUnexpectedCase')
                    job['ExpectedSuites'] += 1
                self.assertTrue(coverage.coverage_errors(broken))

    def test_v6_version_substitution_is_rejected(self):
        for version in (1, 2, 3, 4, 5, 7, '6'):
            with self.subTest(version=version):
                broken = self.manifest_for_version(6)
                broken['CoverageVersion'] = version
                self.assertTrue(coverage.coverage_errors(broken))

    def test_v6_runner_snapshot_prevents_a_historical_plan_downgrade(self):
        current = self.manifest_for_version(6)
        self.assertEqual(coverage.runner_coverage_errors(current, 'CoverageVersion=6;'), [])
        downgraded = self.manifest_for_version(5)
        self.assertEqual(coverage.coverage_errors(downgraded), [])
        self.assertTrue(coverage.runner_coverage_errors(downgraded, 'CoverageVersion=6;'))
        self.assertEqual(coverage.runner_coverage_errors(downgraded, 'CoverageVersion=5;'), [])
        for runner in ('', 'CoverageVersion=5;', 'CoverageVersion=7;', 'CoverageVersion=5;CoverageVersion=6;'):
            with self.subTest(runner=runner):
                self.assertTrue(coverage.runner_coverage_errors(current, runner))
        current['CoverageDefinition'] = 'HistoricalCoreSuitesOnly'
        self.assertTrue(coverage.runner_coverage_errors(current, 'CoverageVersion=6;'))

    def test_v6_missing_duplicate_extra_group_and_mode_are_rejected(self):
        for mutation in ('missing', 'duplicate', 'extra', 'mode'):
            with self.subTest(mutation=mutation):
                broken = self.manifest_for_version(6)
                if mutation == 'missing':
                    broken['Jobs'].pop()
                elif mutation == 'duplicate':
                    broken['Jobs'].append(copy.deepcopy(broken['Jobs'][-1]))
                elif mutation == 'extra':
                    broken['Jobs'].append(dict(Group='UnexpectedGroup', Mode='No', ExpectedSuites=1, Macro=[]))
                else:
                    broken['Jobs'][-1]['Mode'] = 'UnexpectedMode'
                self.assertTrue(coverage.coverage_errors(broken))

    def test_v6_focused_plan_is_still_mandatory(self):
        broken = self.manifest_for_version(6)
        broken['IncludesFocusedAcceptance'] = False
        self.assertTrue(coverage.coverage_errors(broken))

    def test_runner_declares_the_same_v6_feature_list(self):
        text = Path(__file__).with_name('Run-FullOnOffAcceptance.ps1').read_text(encoding='utf-8-sig')
        block = re.search(r'\$currentFeatures\s*=\s*@\((.*?)\)', text, re.S)
        self.assertIsNotNone(block)
        macros = re.findall(r"'([^']+)'", block.group(1))
        self.assertEqual(len(macros), 11)
        self.assertEqual(set(macros), coverage.CONTOUR_ELEMENT_MACROS)
        self.assertRegex(text, r'CoverageVersion\s*=\s*6\b')

    def test_finalizer_uses_dynamic_v6_core_count(self):
        """Проверяет композицию Finalize; первичные верификаторы здесь подставлены."""
        spec = importlib.util.spec_from_file_location('finalizer', Path(__file__).with_name('Finalize-CompleteAcceptance.py'))
        finalizer = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(finalizer)
        source = 'A' * 64
        main = dict(status='PASS', sourceSHA256=source, observedSuites=178, expectedSuites=178,
                    coverageDefinition='AllActiveAcceptanceIncludingFocused', includesFocusedAcceptance=True,
                    jobs=[dict(passed=178)])
        focused = dict(Status='PASS', SourceSHA256=source, IndividualCases=122, PassedAssertions=1)
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            (directory / 'Run.json').write_text('{}', encoding='utf-8')
            (directory / 'Independent.json').write_text(json.dumps(dict(IndependentIntegrals=562, ReviewedNumFailStates=10)), encoding='utf-8')
            def fake_module(name, filename):
                if filename == 'Summarize-FullOnOffAcceptance.py':
                    return SimpleNamespace(summarize=lambda core: main)
                if filename == 'Verify-FocusedAcceptance.py':
                    return SimpleNamespace(verify=lambda path, run: focused)
                self.fail('Unexpected verifier: ' + filename)
            with patch.object(finalizer, 'module', side_effect=fake_module):
                result = finalizer.finalize(directory, directory)
                self.assertEqual(result['CoreSuites'], 178)
                self.assertTrue(result['SingleInvocation'])
                main['observedSuites'] = 177
                with self.assertRaises(ValueError):
                    finalizer.finalize(directory, directory)

    def test_all_current_groups_required(self):
        self.assertEqual(coverage.coverage_errors(self.current), [])
        missing = copy.deepcopy(self.current)
        missing['Jobs'].pop()
        self.assertTrue(coverage.coverage_errors(missing))

    def test_each_feature_is_mandatory(self):
        for macro in coverage.SNAPSHOT_ACCELERATION_MACROS:
            broken = copy.deepcopy(self.current)
            broken['Jobs'][-1]['Macro'].remove(macro)
            broken['Jobs'][-1]['Macro'].append('modTestGeometry.RunGeometryTests')
            self.assertTrue(coverage.coverage_errors(broken), macro)

    def test_duplicate_does_not_replace_missing_feature(self):
        broken = copy.deepcopy(self.current)
        broken['Jobs'][-1]['Macro'][0] = broken['Jobs'][-1]['Macro'][1]
        self.assertTrue(coverage.coverage_errors(broken))

    def test_message_presentation_is_required_in_new_full_runs(self):
        current = copy.deepcopy(self.current)
        current['CoverageVersion'] = 5
        for job in current['Jobs']:
            if job['Group'] == 'CurrentFeatures':
                job.update(ExpectedSuites=6, Macro=sorted(coverage.MESSAGE_PRESENTATION_MACROS))
        self.assertEqual(coverage.coverage_errors(current), [])
        for macro in coverage.MESSAGE_PRESENTATION_MACROS:
            broken = copy.deepcopy(current)
            broken['Jobs'][-1]['Macro'].remove(macro)
            broken['Jobs'][-1]['Macro'].append('modTestCoverage.TestOther')
            self.assertTrue(coverage.coverage_errors(broken), macro)

    def test_focused_plan_cannot_be_omitted(self):
        broken = copy.deepcopy(self.current)
        broken['IncludesFocusedAcceptance'] = False
        self.assertTrue(coverage.coverage_errors(broken))

    def test_snapshot_failure_line_cannot_hide_behind_successful_total(self):
        source = '1' * 64
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / 'CurrentFeatures.txt'
            for failure in ('FAIL: broken snapshot', 'FAIL broken snapshot'):
                path.write_text('\n'.join([
                    'SOURCE_SHA256: ' + source,
                    'SUITE_STARTED: modTestCalculationSnapshot.RunCalculationSnapshotTests; 2026-10-24T00:00:00',
                    failure, 'TOTAL CALCULATION SNAPSHOT: passed=1; failed=0',
                    'SUITE_FINISHED: modTestCalculationSnapshot.RunCalculationSnapshotTests; 2026-10-24T00:00:01',
                    'SOURCE_UNCHANGED: True', 'WATCHDOG_COMPLETED: exit=0; source unchanged=True',
                ]) + '\n', encoding='utf-8')
                result = coverage.load_logs().inspect(path, source, 1)
                self.assertEqual(result['status'], 'INCOMPLETE_OR_FAILED', failure)


if __name__ == '__main__':
    unittest.main()
