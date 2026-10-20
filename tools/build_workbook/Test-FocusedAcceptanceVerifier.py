"""Контроль независимой проверки на аналитическом НДС и невозможной нагрузке."""
import copy
import importlib.util
import sys
import unittest
from pathlib import Path

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT/'tmp/focused-dependencies'))

def module(name,file):
    spec=importlib.util.spec_from_file_location(name,Path(__file__).with_name(file))
    result=importlib.util.module_from_spec(spec);spec.loader.exec_module(result);return result

focused=module('focused','Verify-FocusedAcceptance.py')
full=module('full','Summarize-FullOnOffAcceptance.py')

def analytical_model():
    lines=['CENTER|x=0|y=0','ACCEPTANCE_TOLERANCE|N=0.001|Mx=0.1|My=0.1']
    for x,y in ((-100,-100),(-100,100),(100,-100),(100,100)):
        lines.append(f'FIBER|kind=C|x={x}|y={y}|area=10000')
    lines.append('FIBER|kind=S|x=0|y=0|area=100')
    for mode in ('No','Yes'):
        for kind,points in (('C',((-0.0035,-15.5),(0,0),(0.00015,1.1))),
                            ('S',((-0.025,-350),(0,0),(0.025,350)))):
            flag='True' if mode=='Yes' else 'False'
            lines.append(f'MATERIAL_FLAGS|mode={mode}|profile=PR1|kind={kind}|compressionExtension={flag}|tensionExtension={flag}')
            for strain,stress in points:
                lines.append(f'MATERIAL|mode={mode}|profile=PR1|kind={kind}|strain={strain}|stress={stress}')
    return focused.Model(lines)

class IndependentChecks(unittest.TestCase):
    def test_binary_double_decoder_keeps_all_bits(self):
        import struct
        for value in (-607338.770498685,-223.95934990203004,1e-300,-0.0,1e20):
            encoded='F64:'+struct.pack('>d',value).hex()
            self.assertEqual(struct.pack('>d',focused.number(encoded)),struct.pack('>d',value))
    def test_biaxial_analytical_forces_and_displaced_concrete(self):
        import numpy as np
        actual=analytical_model().integrate(np.array([-0.001,0.000001,-0.000002]),'No','PR1')
        # Eбетона=31000/7; сталь замещает 100 мм2 бетона в центре.
        expected=np.array([-178100.,12400000./7.,-24800000./7.])
        np.testing.assert_allclose(actual,expected,rtol=0,atol=1e-8)

    def test_disabled_extension_clamps_stresses(self):
        import numpy as np
        self.assertAlmostEqual(analytical_model().integrate(np.array([-1.,0.,0.]),'No','PR1')[0],-653450.)

    def test_infinite_extension_is_not_bounded_by_saved_knots(self):
        import numpy as np
        self.assertAlmostEqual(analytical_model().integrate(np.array([-1.,0.,0.]),'Yes','PR1')[0],-178100000.)

    def test_highs_proves_impossible_bounded_load(self):
        import numpy as np
        review=analytical_model().independent_failure(np.array([-178100000.,0.,0.]),'No','PR1')
        self.assertTrue(review['Passed'])
        self.assertEqual(review['Classification'],'ProvenImpossibleByNecessaryForceBounds')
        self.assertGreater(review['StrictMarginN'],0)

    def test_independent_solution_blocks_acceptance_of_numfail(self):
        import numpy as np
        review=analytical_model().independent_failure(np.array([-178100000.,0.,0.]),'Yes','PR1')
        self.assertFalse(review['Passed'])
        self.assertEqual(review['Classification'],'IndependentEquilibriumFoundForNumFail')

class FullScopeChecks(unittest.TestCase):
    def setUp(self):
        self.plan={'CoverageDefinition':'AllActiveAcceptanceIncludingFocused','IncludesFocusedAcceptance':True,
                   'Jobs':[{'Group':g,'Mode':m,'ExpectedSuites':n} for g,n in full.GROUP_COUNTS.items() for m in ('No','Yes')]}

    def test_complete_declared_scope(self):self.assertEqual(full.coverage_errors(self.plan),[])

    def test_omitted_hour_plan_is_rejected(self):
        del self.plan['IncludesFocusedAcceptance'];self.assertTrue(full.coverage_errors(self.plan))

    def test_each_omitted_group_is_rejected(self):
        for index in range(len(self.plan['Jobs'])):
            plan=copy.deepcopy(self.plan);del plan['Jobs'][index];self.assertTrue(full.coverage_errors(plan))

    def test_duplicate_group_cannot_replace_missing_group(self):
        self.plan['Jobs'][-1]=self.plan['Jobs'][0];self.assertTrue(full.coverage_errors(self.plan))

    def test_lowered_suite_count_is_rejected(self):
        self.plan['Jobs'][0]['ExpectedSuites']-=1;self.assertTrue(full.coverage_errors(self.plan))

    def expanded_plan(self):
        return {'CoverageDefinition':'AllActiveAcceptanceIncludingFocused','CoverageVersion':2,'IncludesFocusedAcceptance':True,
                'Jobs':[{'Group':g,'Mode':m,'ExpectedSuites':n} for g,n in full.CURRENT_GROUP_COUNTS.items() for m in ('No','Yes')]}

    def test_expanded_scope_requires_recent_offset_tests(self):
        plan=self.expanded_plan();self.assertEqual(full.coverage_errors(plan),[])
        plan['Jobs']=[j for j in plan['Jobs'] if j['Group']!='SavedOffsets']
        self.assertTrue(full.coverage_errors(plan))

    def test_expanded_geometry_cannot_use_old_count(self):
        plan=self.expanded_plan()
        for job in plan['Jobs']:
            if job['Group']=='Geometry':job['ExpectedSuites']=26
        self.assertTrue(full.coverage_errors(plan))

    def test_unknown_coverage_version_is_rejected(self):
        plan=self.expanded_plan();plan['CoverageVersion']=999
        self.assertTrue(full.coverage_errors(plan))

    def test_individual_rows_require_all_new_suites(self):
        plan=self.expanded_plan();plan['CoverageVersion']=3
        for job in plan['Jobs']:
            if job['Group']=='Geometry':
                job['ExpectedSuites']=36;job['Macro']=[f'Baseline.{i}' for i in range(30)]+list(full.INDIVIDUAL_ROW_MACROS)
        self.assertEqual(full.coverage_errors(plan),[])
        for job in plan['Jobs']:
            if job['Group']=='Geometry':job['Macro'].pop()
        self.assertTrue(full.coverage_errors(plan))

if __name__=='__main__':unittest.main(verbosity=2)
