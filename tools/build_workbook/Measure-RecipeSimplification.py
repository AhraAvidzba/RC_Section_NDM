"""Сравнивает полные рецепты и добавленный общий код по одной методике."""
import argparse
import importlib.util
import json
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('recipe_measure', Path(__file__).with_name('Measure-Audit02Source.py'))
measure = importlib.util.module_from_spec(spec)
spec.loader.exec_module(measure)


def counts(text):
    lines = text.splitlines()
    return {'Lines': len(lines), 'CommentLines': sum(line.lstrip().startswith("'") for line in lines),
            **measure.measure(text)}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--base-ref', required=True)
    parser.add_argument('--report', type=Path, required=True)
    args = parser.parse_args()
    names = ['src/Geometry/CSectionRecipe' + name + '.cls'
             for name in ('Circle', 'RectSet', 'RoundedRectangle', 'HollowRectangle')]
    common = 'src/Geometry/modSectionRecipeRows.bas'
    sources = {}
    for name in names + [common]:
        old = subprocess.run(['git', 'show', args.base_ref + ':' + name], cwd=ROOT, capture_output=True)
        if old.returncode and name != common:
            raise RuntimeError('Baseline source missing: ' + name)
        sources[name] = {'Before': counts(old.stdout.decode('utf-8-sig')),
                         'After': counts((ROOT / name).read_text(encoding='utf-8-sig'))}
    def total(selected, phase):
        return {key: sum(sources[name][phase][key] for name in selected)
                for key in ('Lines', 'CommentLines', 'Instructions', 'Branches', 'Methods')}
    result = {'Baseline': args.base_ref,
              'Method': 'Full files including headers/comments; executable statements from Measure-Audit02Source.measure, with colon statements counted separately',
              'RecipesBefore': total(names, 'Before'), 'RecipesAfter': total(names, 'After'),
              'AddedCommon': sources[common]['After'],
              'TotalBefore': total(names + [common], 'Before'), 'TotalAfter': total(names + [common], 'After'),
              'Sources': sources}
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({key: result[key] for key in ('RecipesBefore', 'RecipesAfter', 'AddedCommon', 'TotalBefore', 'TotalAfter')}, ensure_ascii=False))
