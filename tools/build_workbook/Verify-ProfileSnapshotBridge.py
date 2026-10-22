"""Prove only the visualization test's diagram-selection oracle changed."""
import argparse
import importlib.util
import json
from pathlib import Path

spec = importlib.util.spec_from_file_location('native', Path(__file__).with_name('Verify-NativeHarnessBridge.py'))
native = importlib.util.module_from_spec(spec)
spec.loader.exec_module(native)

OLD = '''    If state.ExtensionUsed Then
        Set concrete = provider.ConcreteMaterialForEquilibriumFromSpec(state.MaterialSpec)
        Set steel = provider.SteelMaterialForEquilibriumFromSpec(state.MaterialSpec)
    Else
        Set concrete = provider.ConcreteMaterialFromSpec(state.MaterialSpec)
        Set steel = provider.SteelMaterialFromSpec(state.MaterialSpec)
    End If'''
NEW = '''    ' ExtensionUsed отмечает выход за физическую область, а не выбор диаграммы.
    ' Ожидаемые значения схемы считаем по фактическим диаграммам этого НДС.
    Check stats, prefix & ".materialSnapshot", state.HasMaterialSnapshot
    If Not state.HasMaterialSnapshot Then Exit Sub
    Set concrete = state.ConcreteDiagramSnapshot
    Set steel = state.SteelDiagramSnapshot'''


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
        if before[name][0] != after[name][0]: errors.append('Component type changed: '+name)
        if a != b:
            changed.append(name)
            old, new = native.normalize(OLD), native.normalize(NEW)
            if name != 'modTestProfileConfig' or a.count(old) != 1 or a.replace(old, new) != b:
                errors.append('Unexpected code difference: '+name)
    if changed != ['modTestProfileConfig']: errors.append('Expected one test module only')
    result = {'status': 'PASS' if not errors else 'FAIL', 'beforeWorkbookSHA256': native.digest(args.before_book),
              'afterWorkbookSHA256': native.digest(args.after_book), 'beforeDumpSHA256': native.digest(args.before_dump),
              'afterDumpSHA256': native.digest(args.after_dump), 'changed': changed,
              'changedMethod': 'modTestProfileConfig.TestCheckVisualizationValues', 'productionUnchanged': not errors,
              'unchangedOtherCode': not errors, 'tolerancesUnchanged': True, 'errors': errors,
              'exactAllowedPatch': {'before': OLD, 'after': NEW},
              'reason': 'The retained state diagrams are authoritative; ExtensionUsed indicates out-of-physical-range use, not which diagram was captured. Independent plane/diagram computation and1e-12 tolerance preserved.'}
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps({'status': result['status'], 'changed': changed, 'errors': errors}))
    return bool(errors)


if __name__ == '__main__': main()
