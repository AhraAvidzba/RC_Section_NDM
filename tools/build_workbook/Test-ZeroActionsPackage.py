"""ДЛЯ ТЕСТОВ: сохранность пользовательской книги при изменении фильтра и справки."""
import argparse
from copy import copy
import importlib.util
import json
from pathlib import Path
from openpyxl import load_workbook

# -------------------- ДЛЯ ТЕСТОВ --------------------
def test_preserved(source, candidate):
    """Сравнивает ввод и сохранённые Results; допускает только новое описание фильтра N."""
    old = load_workbook(source, data_only=False)
    new = load_workbook(candidate, data_only=False)
    try:
        settings = old.defined_names['rngSystemSettings']
        sheet, bounds = next(settings.destinations)
        row = next(row[0].row for row in old[sheet][bounds]
                   if row[0].value == 'Calculation.ZeroAxialForcePerDepth')
        allowed = (sheet, old[sheet].cell(row, 4).coordinate)
        count = 0
        for name in ('Config', 'Сечения', 'Расчет', 'Results'):
            before, after = old[name], new[name]
            cells = set(before._cells) | set(after._cells)
            for position in cells:
                a, b = before.cell(*position), after.cell(*position)
                if (name, a.coordinate) == allowed:
                    assert 'max(hx, hy)' in b.value and 'максимальному' in b.value
                    continue
                assert (a.value, a.data_type) == (b.value, b.data_type), (name, a.coordinate, 'value')
                if name != 'Results':
                    for attribute in ('font', 'fill', 'border', 'alignment', 'protection'):
                        assert copy(getattr(a, attribute)) == copy(getattr(b, attribute)), (name, a.coordinate, attribute)
                    assert a.number_format == b.number_format, (name, a.coordinate, 'number format')
                count += 1
            if name != 'Results':
                assert str(before.merged_cells) == str(after.merged_cells), (name, 'merged cells')
                for position in set(before.row_dimensions) | set(after.row_dimensions):
                    a, b = before.row_dimensions[position], after.row_dimensions[position]
                    assert (a.height, a.hidden, a.outlineLevel) == (b.height, b.hidden, b.outlineLevel), (name, 'row', position)
                for position in set(before.column_dimensions) | set(after.column_dimensions):
                    a, b = before.column_dimensions[position], after.column_dimensions[position]
                    assert (a.width, a.hidden, a.outlineLevel) == (b.width, b.hidden, b.outlineLevel), (name, 'column', position)
        assert {key: item.attr_text for key, item in old.defined_names.items()} == {
            key: item.attr_text for key, item in new.defined_names.items()}, 'Defined names changed'
        return {'ComparedCells': count, 'ExpectedChangedComment': '!'.join(allowed),
                'InputValuesFormulasStylesSizesPreserved': True, 'SavedResultsPreserved': True,
                'DefinedNamesPreserved': True}
    finally:
        old.close()
        new.close()

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ('source', 'candidate', 'clean', 'report'):
        parser.add_argument('--' + name, type=Path, required=True)
    args = parser.parse_args()
    spec = importlib.util.spec_from_file_location('display', Path(__file__).with_name('Test-CalculationDisplayPackage.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    result = {'Status': 'PASS', 'UserWorkbook': test_preserved(args.source, args.candidate),
              'CandidateDisplay': module.test_package(args.candidate), 'CleanDisplay': module.test_package(args.clean)}
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({key: value for key, value in result.items() if key in ('Status', 'UserWorkbook')}, ensure_ascii=False))
