"""Проверяет, что новая обязательная приёмка не теряется при сводном PASS."""
import copy
import importlib.util
from pathlib import Path
import tempfile
import unittest

spec = importlib.util.spec_from_file_location('coverage', Path(__file__).with_name('Summarize-FullOnOffAcceptance.py'))
coverage = importlib.util.module_from_spec(spec)
spec.loader.exec_module(coverage)


class TestFullCoverage(unittest.TestCase):
    """Исторический состав сохраняется; новый требует все три дополнительных пути."""

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
