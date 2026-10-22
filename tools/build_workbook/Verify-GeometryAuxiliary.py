"""Verify that both private acceptance books add only the same test helper."""
import argparse
import importlib.util
import json
from pathlib import Path

spec = importlib.util.spec_from_file_location('bridge', Path(__file__).with_name('Verify-NativeHarnessBridge.py'))
bridge = importlib.util.module_from_spec(spec)
spec.loader.exec_module(bridge)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory', type=Path, required=True)
    parser.add_argument('--book-prefix', default='Auxiliary')
    parser.add_argument('--helper-source', type=Path)
    parser.add_argument('--report', default='AuxiliaryInputsEvidence.json')
    args = parser.parse_args()
    root = args.directory
    helper = args.helper_source or Path(__file__).resolve().parents[2]/'docs/geometry_refactor/examples/modTestFinalGeometryChecks.bas'
    expected_helper = bridge.normalize(helper.read_text(encoding='utf-8-sig'))
    errors, inputs = [], []
    for kind, reference in (('Baseline', 'Baseline/VBA_All_Code.txt'), ('Final', 'VBA_Final.txt')):
        base = bridge.read(root/reference)
        dump = root/f'VBA_{args.book_prefix}{kind}.txt'
        actual = bridge.read(dump)
        added = sorted(set(actual)-set(base))
        changed = sorted(name for name in base if name not in actual or base[name][0] != actual[name][0]
                         or bridge.normalize(base[name][1]) != bridge.normalize(actual[name][1]))
        helper_equal = 'modTestFinalGeometryChecks' in actual and bridge.normalize(actual['modTestFinalGeometryChecks'][1]) == expected_helper
        if len(base) != 149 or len(actual) != 150 or added != ['modTestFinalGeometryChecks'] or changed or not helper_equal:
            errors.append('Unexpected auxiliary code difference: ' + kind)
        inputs.append({'kind': kind, 'book': f'{args.book_prefix}{kind}.xlsm',
                       'bookSHA256': bridge.digest(root/f'{args.book_prefix}{kind}.xlsm'),
                       'dump': dump.name, 'dumpSHA256': bridge.digest(dump),
                       'referenceDump': reference, 'referenceDumpSHA256': bridge.digest(root/reference),
                       'components': len(actual), 'added': added, 'changed': changed, 'helperEqual': helper_equal})
    result = {'status': 'PASS' if not errors else 'FAIL', 'inputs': inputs,
              'helperSource': str(helper), 'helperSourceSHA256': bridge.digest(helper), 'errors': errors,
              'scope': 'All 149 original components unchanged; one identical private helper only; no helper in production release'}
    (root/args.report).write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps({'status': result['status'], 'errors': errors}))
    return bool(errors)


if __name__ == '__main__':
    raise SystemExit(main())
