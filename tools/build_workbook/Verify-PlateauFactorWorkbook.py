"""Проверяет только согласованные изменения наклона; остальной ввод и Results точные."""
import argparse
import copy
import hashlib
import importlib.util
import json
import re
from pathlib import Path

spec = importlib.util.spec_from_file_location('reader', Path(__file__).with_name('Verify-SectionRecipeGuide.py'))
reader = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reader)
S = reader.S
OLD = 'Solver.PlateauSlope.Enabled'
NEW = 'General.PlateauSlopeFactor'


def col(address):
    return re.sub(r'\d+', '', address)


def sorted_children(element):
    if element is None:
        return None
    result = copy.deepcopy(element)
    for item in result.iter():
        # Excel создаёт новые идентификаторы при копировании. Это служебная
        # история редактирования, а не назначение ссылки или проверки ввода.
        for key in list(item.attrib):
            if key.endswith('}uid'):
                del item.attrib[key]
    result[:] = sorted(result, key=lambda item: str(sorted(item.attrib.items())))
    return reader.canonical(result)


def expand(reference):
    def index(letters):
        result = 0
        for letter in letters:
            result = result * 26 + ord(letter) - 64
        return result
    def letters(index):
        result = ''
        while index:
            index, remainder = divmod(index - 1, 26)
            result = chr(65 + remainder) + result
        return result
    for part in reference.split():
        match = re.fullmatch(r'([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?', part.replace('$', ''))
        if not match:
            raise ValueError('Unsupported range: ' + part)
        left, top = index(match[1]), int(match[2])
        right, bottom = index(match[3] or match[1]), int(match[4] or match[2])
        if (right-left+1) * (bottom-top+1) > 200000:
            raise ValueError('Validation range is unexpectedly large')
        for row in range(top, bottom+1):
            for column in range(left, right+1):
                yield letters(column) + str(row)


def validations(sheet):
    result = {}
    parent = sheet['document'].find(S + 'dataValidations')
    if parent is not None:
        for item in parent:
            value = copy.deepcopy(item)
            reference = value.attrib.pop('sqref')
            for key in list(value.attrib):
                if key.endswith('}uid'):
                    del value.attrib[key]
            for address in expand(reference):
                if address in result:
                    raise ValueError('Duplicate validation: ' + address)
                result[address] = reader.canonical(value)
    return result


def verify(before, after):
    old, new = reader.read(before), reader.read(after)
    errors, checked = [], 0
    if set(old['sheets']) != set(new['sheets']):
        errors.append('Sheet inventory changed')
    if old['names'] != new['names']:
        errors.append('Named ranges changed')
    old_config, new_config = old['sheets']['Config'], new['sheets']['Config']
    bounds = old['names']['rngSystemSettings'].rsplit('!', 1)[1]
    system_rows = sorted({reader.row_of(address) for address in expand(bounds)})
    previous = [old_config['cells'].get('A'+str(row), {}).get('value') for row in system_rows]
    following = [new_config['cells'].get('A'+str(row), {}).get('value') for row in system_rows]
    found = previous.index(OLD)
    expected = previous.copy()
    expected.pop(found)
    expected.insert(expected.index('General.DiagramExtension'), NEW)
    if following != expected:
        errors.append('Registry differs from the single authorized row relocation')
    source_rows = list(system_rows)
    moved = source_rows.pop(found)
    source_rows.insert(expected.index(NEW), moved)
    mapping = dict(zip(source_rows, system_rows))
    factor_row = system_rows[following.index(NEW)]
    factor = 'B' + str(factor_row)
    if new_config['cells'][factor]['formula'] is not None or float(new_config['cells'][factor]['value']) != 0.001:
        errors.append('Default is not numeric 0.001')
    old_help, new_help = old['sheets']['Справка'], new['sheets']['Справка']
    anchor = next(reader.row_of(a) for a,c in old_help['cells'].items() if c['value'] == OLD)
    end = anchor
    while old_help['cells'].get('A'+str(end+1), {}).get('value'):
        end += 1
    help_text = ' '.join(new_help['cells'].get('A'+str(row), {}).get('value') or '' for row in range(anchor, end+1))
    for text in (NEW, '0,001·E', '0,1%', 'α = 0', 'не нормативный коэффициент', 'После изменения коэффициента'):
        if text not in help_text:
            errors.append('Help explanation missing: ' + text)
    if OLD in help_text:
        errors.append('Old Boolean remains in Help')
    controls = next(reader.row_of(a) for a,c in old_config['cells'].items() if col(a) == 'AI' and (c['value'] or '').startswith('Контрольные точки диаграмм'))
    def mapped(address):
        row = reader.row_of(address)
        return col(address) + str(mapping[row]) if col(address) in ('A','B','C','D','E') and row in mapping else address
    def style_without_border(style):
        return style[0], style[1][:2] + style[1][3:], style[2]
    def effective_border(sheet, address, edge):
        column, row = col(address), reader.row_of(address)
        merge_parent = sheet['document'].find(S+'mergeCells')
        if merge_parent is not None:
            for merge in merge_parent:
                first, last = merge.get('ref').split(':')
                if reader.row_of(first) <= row <= reader.row_of(last) and col(first) <= column <= col(last):
                    if edge == 'left' and column != col(first) or edge == 'right' and column != col(last):
                        return None
        def edge_value(at, side):
            cell = sheet['cells'].get(at)
            if cell is None:
                return None
            border = cell['style'][1][2]
            value = next(item for item in border[3] if item[0] == S+side)
            line = dict(value[1]).get('style')
            if line is None:
                return None
            color = dict(value[3][0][1]) if value[3] else {}
            if color in ({}, {'indexed':'64'}, {'auto':'1'}, {'rgb':'FF000000'}):
                color = {'rgb':'FF000000'}
            return line, tuple(sorted(color.items()))
        value = edge_value(address, edge)
        if value is None and edge in ('top','bottom'):
            value = edge_value(column+str(row+(-1 if edge == 'top' else 1)), 'bottom' if edge == 'top' else 'top')
        return value
    def same_style(sheet, source, target, a, b):
        if a == b:
            return True
        if style_without_border(a) != style_without_border(b):
            return False
        return all(effective_border(old['sheets'][sheet], source, edge) == effective_border(new['sheets'][sheet], target, edge)
                   for edge in ('left','right','top','bottom'))
    for sheet, data in old['sheets'].items():
        actual = new['sheets'][sheet]
        for address, cell in data['cells'].items():
            target = mapped(address) if sheet == 'Config' else address
            observed = actual['cells'].get(target)
            checked += 1
            allowed_factor = sheet == 'Config' and target in ('A'+str(factor_row), factor, 'D'+str(factor_row))
            allowed_help = sheet == 'Справка' and col(address) in ('A','B','C','D','E','F') and anchor <= reader.row_of(address) <= end
            allowed_control = sheet == 'Config' and col(address) in ('AM','AN') and controls+2 <= reader.row_of(address) <= controls+53
            if allowed_factor or allowed_help or allowed_control:
                if observed is None:
                    errors.append(sheet+'!'+target+': authorized cell disappeared')
                elif target != factor and not same_style(sheet, address, target, cell['style'], observed['style']):
                    errors.append(sheet+'!'+target+': style changed')
                if allowed_help and col(address) != 'A' and observed and (cell['value'],cell['formula']) != (observed['value'],observed['formula']):
                    errors.append(sheet+'!'+target+': unexpected text beside merged Help')
                continue
            expected_cell = copy.deepcopy(cell)
            expected_cell['attrs']['r'] = target
            if observed and same_style(sheet, address, target, cell['style'], observed['style']):
                expected_cell['style'] = observed['style']
            if expected_cell != observed:
                errors.append(sheet+'!'+target+': user cell/formula/style changed')
        expected_addresses = {mapped(a) if sheet == 'Config' else a for a in data['cells']}
        for address, cell in actual['cells'].items():
            if address not in expected_addresses and (cell['value'] or cell['formula']):
                errors.append(sheet+'!'+address+': unexpected new data')
        for row, attrs in data['rows'].items():
            observed = actual['rows'].get(row)
            if sheet == 'Справка' and anchor < row <= end:
                attrs = {k:v for k,v in attrs.items() if k not in ('ht','customHeight')}
                observed = {k:v for k,v in (observed or {}).items() if k not in ('ht','customHeight')}
            if attrs != observed:
                errors.append(sheet+': row metadata changed '+str(row))
        if reader.canonical(data['document'].find(S+'cols')) != reader.canonical(actual['document'].find(S+'cols')):
            errors.append(sheet+': column widths changed')
        old_merges = copy.deepcopy(data['document'].find(S+'mergeCells'))
        if sheet == 'Config' and old_merges is not None:
            for item in old_merges:
                item.set('ref', ':'.join(mapped(a) for a in item.get('ref').split(':')))
        if sorted_children(old_merges) != sorted_children(actual['document'].find(S+'mergeCells')):
            errors.append(sheet+': merge inventory changed')
        old_links = copy.deepcopy(data['document'].find(S+'hyperlinks'))
        if sheet == 'Config' and old_links is not None:
            for item in old_links:
                item.set('ref', mapped(item.get('ref')))
        if sorted_children(old_links) != sorted_children(actual['document'].find(S+'hyperlinks')):
            errors.append(sheet+': hyperlinks changed')
        prior_validations, observed_validations = validations(data), validations(actual)
        if sheet == 'Config':
            prior_validations = {mapped(a):v for a,v in prior_validations.items()}
            prior_validations.pop(factor, None)
            numeric = observed_validations.pop(factor, None)
            if numeric is None or dict(numeric[1]).get('type') != 'decimal' or dict(numeric[1]).get('operator') != 'greaterThanOrEqual':
                errors.append('Factor validation is not decimal >= 0')
        if prior_validations != observed_validations:
            errors.append(sheet+': other validation sources or rules changed')
    if old['drawings'] != new['drawings']:
        errors.append('Drawing geometry changed')
    return {'AllPassed': not errors, 'UserCellsChecked': checked, 'FactorAddress': factor,
            'DefaultFactor': 0.001, 'HelpRows': [anchor,end], 'RowMapping': mapping,
            'BeforeSHA256': hashlib.sha256(before.read_bytes()).hexdigest(),
            'WorkbookSHA256': hashlib.sha256(after.read_bytes()).hexdigest(), 'Errors': errors}


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--before', type=Path, required=True)
    parser.add_argument('--book', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    args = parser.parse_args()
    report = verify(args.before, args.book)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k:v for k,v in report.items() if k != 'RowMapping'}, ensure_ascii=False))
    raise SystemExit(0 if report['AllPassed'] else 1)
