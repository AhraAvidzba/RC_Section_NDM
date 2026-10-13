"""Prove the sole change is a configured-tolerance assertion plus non-overshoot."""
import argparse
import importlib.util
import json
import re
from pathlib import Path

spec = importlib.util.spec_from_file_location('native', Path(__file__).with_name('Verify-NativeHarnessBridge.py'))
native = importlib.util.module_from_spec(spec)
spec.loader.exec_module(native)

OLD = '''    CheckClose stats, prefix & ".Capacity.ToleranceLambda.sameLimit", changed.LambdaCapacity, baseline.LambdaCapacity, 0.002'''
NEW = '''    ' Грубая бисекция возвращает допустимую нижнюю сторону интервала.
    ' Отличие ограничено заданным lambda-допуском плюс точность базового поиска.
    CheckClose stats, prefix & ".Capacity.ToleranceLambda.sameLimit", changed.LambdaCapacity, baseline.LambdaCapacity, _
        CDbl(ValueCell(table, "Capacity.ToleranceLambda", 2).Value2) + 0.000001
    Check stats, prefix & ".Capacity.ToleranceLambda.conservative", changed.LambdaCapacity <= baseline.LambdaCapacity + 0.000001'''


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ('before-book', 'after-book', 'before-dump', 'after-dump', 'report'):
        parser.add_argument('--'+name, type=Path, required=True)
    args = parser.parse_args()
    before, after = native.read(args.before_dump), native.read(args.after_dump)
    errors, changed = [], []
    if set(before) != set(after) or len(after) != 149:
        errors.append('Expected the same 149 components')
    for name in sorted(set(before) & set(after)):
        a, b = native.normalize(before[name][1]), native.normalize(after[name][1])
        if before[name][0] != after[name][0]:
            errors.append('Component type changed: '+name)
        if a != b:
            changed.append(name)
            old, new = native.normalize(OLD), native.normalize(NEW)
            if name != 'modTestConfiguration' or a.count(old) != 1 or a.replace(old, new) != b:
                errors.append('Unexpected code difference: '+name)
    if changed != ['modTestConfiguration']:
        errors.append('Expected exactly one changed test module')
    body = after['modTestConfiguration'][1]
    block = re.search(r'(?ms)^Private Sub CheckCapacityNumericEffects\b.*?^End Sub\b', body)
    if not block or native.normalize(NEW) not in native.normalize(block.group()):
        errors.append('Change is not isolated in the intended private test')
    result = {'status': 'PASS' if not errors else 'FAIL', 'beforeWorkbookSHA256': native.digest(args.before_book),
              'afterWorkbookSHA256': native.digest(args.after_book), 'beforeDumpSHA256': native.digest(args.before_dump),
              'afterDumpSHA256': native.digest(args.after_dump), 'changed': changed,
              'changedMethod': 'modTestConfiguration.CheckCapacityNumericEffects',
              'productionUnchanged': not errors, 'unchangedOtherCode': not errors, 'errors': errors,
              'exactAllowedPatch': {'before': OLD, 'after': NEW},
              'engineeringReason': 'Absolute coarse lambda error is bounded by the selected 0.02 tolerance plus baseline1e-6; equilibrium/physical checks and linear-fixture0.002 remain unchanged'}
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps({'status': result['status'], 'changed': changed, 'errors': errors}))
    return bool(errors)


if __name__ == '__main__':
    raise SystemExit(main())
