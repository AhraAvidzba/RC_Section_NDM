"""ДЛЯ ТЕСТОВ: строгий разбор итогов кручения и отказов в полных первичных журналах."""
import importlib.util
from pathlib import Path
import tempfile
import unittest

spec=importlib.util.spec_from_file_location('full_logs',Path(__file__).with_name('Summarize-GeometryFull.py'))
logs=importlib.util.module_from_spec(spec);spec.loader.exec_module(logs)

class FullLogTotalsTests(unittest.TestCase):
    # -------------------- ДЛЯ ТЕСТОВ --------------------
    def test_inspect(self):
        """Каждый предусмотренный набор принимает только свой единственный и подтверждённый итог."""
        for name,label in logs.TORSION_TOTAL_LABELS.items():
            with self.subTest(name=name):
                result=self.read_test_log(name,f'PASS first\nPASS second\n{label}: Passed=2; Failed=0\n')
                self.assertEqual(result['status'],'PASS');self.assertEqual(result['suites'][0]['passed'],2)

    # -------------------- ДЛЯ ТЕСТОВ --------------------
    def test_rejections(self):
        """Не принимает отсутствующие/чужие/повторные итоги, неверные счётчики и скрытые ошибки."""
        name='modTestTorsionPost.RunTorsionPostTests'
        cases={
            'missing':'PASS first\n',
            'wrong_label':'PASS first\nTorsionPostUnits: Passed=1; Failed=0\n',
            'duplicate':'PASS first\nTorsionPost: Passed=1; Failed=0\nTorsionPost: Passed=1; Failed=0\n',
            'foreign_total':'PASS first\nTorsionPost: Passed=1; Failed=0\nTorsionPostUnits: Passed=1; Failed=0\n',
            'zero_assertions':'TorsionPost: Passed=0; Failed=0\n',
            'wrong_count':'PASS first\nTorsionPost: Passed=2; Failed=0\n',
            'failed_total':'PASS first\nFAIL second\nTorsionPost: Passed=1; Failed=1\n',
            'hidden_fail':'PASS first\nFAIL second\nTorsionPost: Passed=1; Failed=0\n',
            'nested_fail':'PASS first\nTOTAL_INNER: passed=0; failed=1\nTorsionPost: Passed=1; Failed=0\n',
            'runtime':'PASS first\nRUNTIME ERROR: stopped\nTorsionPost: Passed=1; Failed=0\n',
        }
        for reason,body in cases.items():
            with self.subTest(reason=reason):self.assertNotEqual(self.read_test_log(name,body)['status'],'PASS')

    # -------------------- ДЛЯ ТЕСТОВ --------------------
    def test_source_and_completion(self):
        """Итог не заменяет правильный источник и успешное завершение suite/runner."""
        name='modTestTorsionPost.RunTorsionPostTests';body='PASS first\nTorsionPost: Passed=1; Failed=0\n'
        self.assertNotEqual(self.read_test_log(name,body,finished=False)['status'],'PASS')
        self.assertNotEqual(self.read_test_log(name,body,terminal=False)['status'],'PASS')
        self.assertNotEqual(self.read_test_log(name,body,source='B'*64)['status'],'PASS')

    # -------------------- ДЛЯ ТЕСТОВ --------------------
    def test_generic_totals(self):
        """Обычные TOTAL и запрет вложенных отказов сохраняют свой контракт."""
        name='modTestGeometry.RunGeometryTests'
        self.assertEqual(self.read_test_log(name,'OK: first\nTOTAL: passed=1; failed=0\n')['status'],'PASS')
        self.assertNotEqual(self.read_test_log(name,'TOTAL_INNER: passed=0; failed=1\nTOTAL: passed=1; failed=0\n')['status'],'PASS')

    # -------------------- ДЛЯ ТЕСТОВ --------------------
    def read_test_log(self,name,body,finished=True,terminal=True,source='A'*64):
        """Создаёт изолированный минимальный журнал с реальным форматом границ suite."""
        text=f'SOURCE_SHA256: {source}\nSUITE_STARTED: {name}; 2026-11-20T01:00:00\n'+body
        if finished:text+=f'SUITE_FINISHED: {name}; 2026-11-20T01:00:01\n'
        if terminal:text+='SOURCE_UNCHANGED: True\nWATCHDOG_COMPLETED: exit=0; source unchanged=True\n'
        with tempfile.TemporaryDirectory() as temporary:
            path=Path(temporary)/'Primary.txt';path.write_text(text,encoding='utf-8')
            return logs.inspect(path,'a'*64,1)

if __name__=='__main__':unittest.main(verbosity=2)
