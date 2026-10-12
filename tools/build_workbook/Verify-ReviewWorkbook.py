"""Независимая XML-проверка новой Справки и сохранения пользовательских данных."""
import argparse
import copy
import hashlib
import importlib.util
import json
import re
from collections import Counter
from pathlib import Path

spec = importlib.util.spec_from_file_location('xmlreader', Path(__file__).with_name('Verify-SectionRecipeGuide.py'))
xmlreader = importlib.util.module_from_spec(spec)
spec.loader.exec_module(xmlreader)
S = xmlreader.S
TEXT_CHANGES=json.loads(Path(__file__).with_name('StandaloneHelpText.json').read_text(encoding='utf-8'))
MANUALS=json.loads(Path(__file__).with_name('NormativeHelpSources.json').read_text(encoding='utf-8'))

def canonical_merges(node):
    # Excel меняет порядок записей объединений после сдвига строк на другом листе.
    # Сохраняем строгую проверку всех диапазонов, повторов и атрибутов.
    if node is None: return xmlreader.canonical(None)
    ordered=copy.deepcopy(node)
    ordered[:]=sorted(ordered, key=lambda child: child.get('ref',''))
    return xmlreader.canonical(ordered)

def owned_range(book):
    match = re.fullmatch(r"'?Справка'?!\$A\$(\d+):\$F\$(\d+)", book['names']['rngSectionRecipeGuide'])
    if not match: raise ValueError('Invalid guide owner range')
    return tuple(map(int, match.groups()))

def rows_except_guide(book):
    first, last = owned_range(book)
    cells = book['sheets']['Справка']['cells']
    byrow = {}
    for address, cell in cells.items():
        row = xmlreader.row_of(address)
        if first <= row <= last: continue
        if not cell['value'] and not cell['formula']: continue
        byrow.setdefault(row, {})[re.sub(r'\d+', '', address)] = (TEXT_CHANGES.get(cell['value'],cell['value']), cell['formula'])
    authorized={tuple((column,value,None) for column,value in (('A',manual['caption']),('B',manual['description']))) for manual in MANUALS if manual['kind']=='manual'}
    authorized.add((('A','Важно',None),('B','Справка описывает работу программы и смысл ее настроек. Если применение пункта СП еще не проверено, это прямо указано в соответствующем пояснении.',None)))
    rows=Counter(tuple((column, *value) for column, value in sorted(data.items())) for data in byrow.values())
    for row in authorized:
        if rows[row]: rows[row]-=1
    return +rows

def verify(before, after, layout, halve_default_rebar_offset=False):
    old, new = xmlreader.read(before), xmlreader.read(after)
    content = json.loads(layout.read_text(encoding='utf-8-sig'))
    errors, cells_checked, caches = [], 0, []
    authorized = {}
    if halve_default_rebar_offset:
        # Единственное числовое исключение — явно согласованный default 60 -> 30.
        # Адрес получаем из настоящего named range, а не из положения Config.
        ref = old['names']['rngPlotAnnotationSettings']
        sheet, bounds = ref.rsplit('!', 1)
        if sheet.strip("'=") != 'Config': raise ValueError('Annotation settings are outside Config')
        match = re.fullmatch(r'\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)', bounds)
        if not match: raise ValueError('Invalid annotation settings range')
        column = 0
        for char in match[1]: column = column * 26 + ord(char) - 64
        def letters(number):
            result = ''
            while number:
                number, remainder = divmod(number - 1, 26)
                result = chr(65 + remainder) + result
            return result
        for row in range(int(match[2]), int(match[4]) + 1):
            if old['sheets']['Config']['cells'].get(letters(column)+str(row), {}).get('value') == 'Offset':
                address = letters(column + 1) + str(row)
                if old['sheets']['Config']['cells'][address]['value'] == '60': authorized[('Config', address)] = '30'
                break
        else: raise ValueError('Offset row is missing')
    if set(old['sheets']) != set(new['sheets']): errors.append('Worksheet inventory changed')
    for sheet, data in old['sheets'].items():
        if sheet == 'Справка': continue
        observed = new['sheets'][sheet]['cells']
        for address, cell in data['cells'].items():
            actual = observed.get(address)
            if actual is None:
                if cell['value'] or cell['formula']: errors.append(sheet+'!'+address+': missing data')
                continue
            cells_checked += 1
            if cell['formula'] != actual['formula']: errors.append(sheet+'!'+address+': formula changed')
            elif authorized.get((sheet, address), TEXT_CHANGES.get(cell['value'],cell['value'])) != actual['value']:
                if cell['formula']: caches.append(sheet+'!'+address)
                else: errors.append(sheet+'!'+address+': value changed')
            if cell['style'] != actual['style']: errors.append(sheet+'!'+address+': style changed')
        for address, cell in observed.items():
            if address not in data['cells'] and (cell['value'] or cell['formula']): errors.append(sheet+'!'+address+': new data')
        # Списки, размеры столбцов, объединения и ссылки проверяются отдельно.
        for tag in ('cols', 'dataValidations'):
            if xmlreader.canonical(data['document'].find(S+tag)) != xmlreader.canonical(new['sheets'][sheet]['document'].find(S+tag)):
                errors.append(sheet+': '+tag+' changed')
        if canonical_merges(data['document'].find(S+'mergeCells')) != canonical_merges(new['sheets'][sheet]['document'].find(S+'mergeCells')):
            errors.append(sheet+': mergeCells changed')
        old_links = data['document'].find(S+'hyperlinks')
        new_links = new['sheets'][sheet]['document'].find(S+'hyperlinks')
        if old_links is not None:
            observed_links = {link.get('ref'): link for link in new_links} if new_links is not None else {}
            for link in old_links:
                actual = observed_links.get(link.get('ref'))
                if actual is None: errors.append(sheet+': hyperlink missing'); continue
                old_target, new_target = link.get('location',''), actual.get('location','')
                if 'Справка' in old_target:
                    source_address = old_target.rsplit('!',1)[-1].replace('$','')
                    target_address = new_target.rsplit('!',1)[-1].replace('$','')
                    original = old['sheets']['Справка']['cells'].get(source_address,{}).get('value')
                    moved = new['sheets']['Справка']['cells'].get(target_address,{}).get('value')
                    if original != moved: errors.append(sheet+'!'+link.get('ref')+': link points to another setting')
                elif xmlreader.canonical(link) != xmlreader.canonical(actual): errors.append(sheet+': unrelated hyperlink changed')
    for name, value in old['names'].items():
        if name == 'rngSectionRecipeGuide': continue
        if new['names'].get(name) != value: errors.append('Named range changed: '+name)
    if set(old['names']) != set(new['names']): errors.append('Named range inventory changed')
    if rows_except_guide(old) != rows_except_guide(new): errors.append('Original Help text/data outside owned guide changed')
    first, last = owned_range(new)
    if last-first+1 != len(content['rows']): errors.append('Brief guide length differs')
    help_cells = new['sheets']['Справка']['cells']
    norm_rows=[xmlreader.row_of(address) for address,cell in help_cells.items() if re.fullmatch(r'A\d+',address) and cell['value']=='СП 63.13330.2018']
    if len(norm_rows)!=1: errors.append('SP63 normative source row is missing or repeated')
    else:
        for offset,manual in enumerate(MANUALS):
            row=norm_rows[0]+offset
            if help_cells.get('A'+str(row),{}).get('value')!=manual['caption'] or help_cells.get('B'+str(row),{}).get('value')!=manual['description']:
                errors.append('Manual normative source differs from approved text')
        if help_cells.get('A'+str(norm_rows[0]+len(MANUALS)),{}).get('value')=='Важно': errors.append('Removed Important row remains')
    for index, row in enumerate(content['rows']):
        if help_cells.get('A'+str(first+index),{}).get('value') != row['text']: errors.append('Brief guide text differs at '+str(index))
    titles = [row['value'] for address,row in sorted(help_cells.items(),key=lambda pair: xmlreader.row_of(pair[0]))
              if re.fullmatch(r'A\d+',address) and row['value'] in (
                  'Справка по результатам расчета листа Results','Добавление нового типа сечения','Справка по настройкам листа Config')]
    if titles != ['Справка по результатам расчета листа Results','Добавление нового типа сечения','Справка по настройкам листа Config']:
        errors.append('Main final Help order differs')
    result = {'status': 'REVIEW_XML_PRESERVATION_PASS' if not errors else 'FAIL', 'userCellsChecked': cells_checked,
              'beforeSHA256': hashlib.sha256(before.read_bytes()).hexdigest(),
              'afterSHA256': hashlib.sha256(after.read_bytes()).hexdigest(),
              'formulaCacheRefreshes': caches, 'helpOriginalTextPreserved': rows_except_guide(old) == rows_except_guide(new),
              'guideRows': last-first+1, 'guideFirst': first, 'guideLast': last,
              'authorizedSettingChanges': [{'sheet': sheet, 'address': address, 'before': '60', 'after': value,
                                           'reason': 'User requested half of the default rebar label Offset'}
                                          for (sheet, address), value in authorized.items()], 'errors': errors}
    return result

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--before',type=Path,required=True)
    parser.add_argument('--after',type=Path,required=True)
    parser.add_argument('--layout',type=Path,default=Path('docs/geometry_refactor/SECTION_RECIPE_GUIDE_LAYOUT.json'))
    parser.add_argument('--report',type=Path,required=True)
    parser.add_argument('--halve-default-rebar-offset', action='store_true')
    args=parser.parse_args()
    result=verify(args.before,args.after,args.layout,args.halve_default_rebar_offset)
    args.report.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False))
    if result['errors']: raise SystemExit(1)

if __name__=='__main__':main()
