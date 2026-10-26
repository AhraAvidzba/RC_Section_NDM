"""ДЛЯ ТЕСТОВ: восстанавливает точное оформление входов в отдельном ZIP-пакете XLSM.

Исходные книги не изменяются. VBA, новые справка/подписи, формулы и значения
берутся из подготовленного пакета; стили и rich text неизменённых входных
ячеек — из контрольной книги. Числовые ID стилей между книгами не копируются.
"""
from __future__ import annotations
import argparse
import copy
import hashlib
import importlib.util
import json
import zipfile
from pathlib import Path
from lxml import etree as X

HERE = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location('preservation', HERE / 'TestVerifyOpenXmlPreservation.py')
VERIFY = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(VERIFY)
NS = VERIFY.NS
S = '{' + NS['s'] + '}'
INPUTS = ('Config', 'Сечения', 'Расчет', 'Расчёт')


def signature(node):
    """Сравнивает XML-структуру без зависимости от префиксов и порядка атрибутов."""
    return (node.tag, tuple(sorted(node.attrib.items())), node.text or '',
            tuple(signature(child) for child in node))


def serialized(node):
    """lxml сохраняет исходные nsmap и mc:Ignorable; форматирование не пересобирается."""
    return X.tostring(node, encoding='UTF-8', xml_declaration=True, standalone=True)


class StyleMerge:
    """Переносит зависимости XF по содержимому, сохраняя существующие target ID."""
    ORDER = ('numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs',
             'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst')

    def __init__(self, source, target):
        self.source, self.target = source, target
        self.cache = {}
        self.additions = {}

    def collection(self, root, name):
        item = root.find(S + name)
        if item is None:
            item = X.Element(S + name, count='0')
            position = self.ORDER.index(name)
            next_item = next((child for child in root
                              if X.QName(child).localname in self.ORDER and
                              self.ORDER.index(X.QName(child).localname) > position), None)
            if next_item is None:
                root.append(item)
            else:
                root.insert(root.index(next_item), item)
        return item

    def add(self, name, item):
        target = self.collection(self.target, name)
        token = signature(item)
        for index, current in enumerate(target):
            if signature(current) == token:
                return index
        target.append(copy.deepcopy(item))
        target.set('count', str(len(target)))
        self.additions[name] = self.additions.get(name, 0) + 1
        return len(target) - 1

    def number_format(self, number):
        source = self.source.find(S + 'numFmts')
        item = next((n for n in source if int(n.get('numFmtId')) == number), None) if source is not None else None
        if item is None:
            return number  # Встроенный формат задаётся стандартным числом, без записи numFmt.
        target = self.collection(self.target, 'numFmts')
        attributes = {k: v for k, v in item.attrib.items() if k != 'numFmtId'}
        for current in target:
            if {k: v for k, v in current.attrib.items() if k != 'numFmtId'} == attributes:
                return int(current.get('numFmtId'))
        number = max([163] + [int(n.get('numFmtId')) for n in target]) + 1
        added = copy.deepcopy(item)
        added.set('numFmtId', str(number))
        target.append(added)
        target.set('count', str(len(target)))
        self.additions['numFmts'] = self.additions.get('numFmts', 0) + 1
        return number

    def mapped(self, name, index):
        key = (name, index)
        if key in self.cache:
            return self.cache[key]
        source = self.source.find(S + name)
        if source is None or not 0 <= index < len(source):
            raise ValueError('Missing source style dependency: ' + repr(key))
        item = copy.deepcopy(source[index])
        if name in ('cellStyleXfs', 'cellXfs'):
            for attribute, group in (('fontId', 'fonts'), ('fillId', 'fills'), ('borderId', 'borders')):
                if attribute in item.attrib:
                    item.set(attribute, str(self.mapped(group, int(item.get(attribute)))))
            if 'numFmtId' in item.attrib:
                item.set('numFmtId', str(self.number_format(int(item.get('numFmtId')))))
            if 'xfId' in item.attrib:
                item.set('xfId', str(self.mapped('cellStyleXfs', int(item.get('xfId')))))
        result = self.add(name, item)
        self.cache[key] = result
        return result


def shared_item(root, cell):
    if cell.get('t') == 's':
        value = cell.find(S + 'v')
        return root[int(value.text)] if value is not None else None
    if cell.get('t') == 'inlineStr':
        return cell.find(S + 'is')
    return None


def restore(before_path, base_path, output_path, report_path):
    before_path, base_path, output_path = map(Path, (before_path, base_path, output_path))
    if output_path.exists() or output_path.resolve() in (before_path.resolve(), base_path.resolve()):
        raise ValueError('Output must be a new distinct path')
    original_source_hash = VERIFY.digest(before_path.read_bytes())
    original_base_hash = VERIFY.digest(base_path.read_bytes())
    old, base = VERIFY.Book(before_path), VERIFY.Book(base_path)
    if VERIFY.setting_rows(old)['Mesh.StepMode'] != VERIFY.setting_rows(base)['Mesh.StepMode']:
        raise ValueError('Style source must already contain the approved Config row insertion')
    source_styles = X.fromstring(old.parts['xl/styles.xml'])
    target_styles = X.fromstring(base.parts['xl/styles.xml'])
    styles = StyleMerge(source_styles, target_styles)
    source_strings = X.fromstring(old.parts['xl/sharedStrings.xml'])
    target_strings = X.fromstring(base.parts['xl/sharedStrings.xml'])
    string_ids = {signature(n): i for i, n in enumerate(target_strings)}
    changes, input_reports, text_changes = {}, [], []
    modified_sheets = {}
    for name in INPUTS:
        if name not in old.sheets:
            continue
        source = X.fromstring(old.parts[old.sheets[name]['part']])
        target = X.fromstring(base.parts[base.sheets[name]['part']])
        source_cells = {n.get('r'): n for n in source.findall('s:sheetData/s:row/s:c', NS)}
        target_cells = {n.get('r'): n for n in target.findall('s:sheetData/s:row/s:c', NS)}
        repaired_styles, repaired_strings, unchanged_text = [], [], 0
        for ref, current in target_cells.items():
            original = source_cells.get(ref)
            if original is None:
                continue
            if old.style_at(name, ref) != base.style_at(name, ref):
                index = styles.mapped('cellXfs', int(original.get('s', '0')))
                current.set('s', str(index))
                repaired_styles.append({'Cell': ref, 'SourceStyle': original.get('s', '0'), 'MappedStyle': index})
            old_value = old.sheets[name]['cells'].get(ref)
            new_value = base.sheets[name]['cells'].get(ref)
            if old_value != new_value:
                text_changes.append({'Sheet': name, 'Cell': ref, 'Before': old_value, 'After': new_value})
                continue  # Новые каталожные описания не откатываются.
            original_string, current_string = shared_item(source_strings, original), shared_item(target_strings, current)
            if original_string is None:
                continue
            unchanged_text += 1
            if current_string is not None and signature(original_string) == signature(current_string):
                continue
            if original.get('t') == 's':
                key = signature(original_string)
                if key not in string_ids:
                    string_ids[key] = len(target_strings)
                    target_strings.append(copy.deepcopy(original_string))
                inline = current.find(S + 'is')
                if inline is not None:
                    current.remove(inline)
                value = current.find(S + 'v')
                if value is None:
                    value = X.SubElement(current, S + 'v')
                current.set('t', 's')
                value.text = str(string_ids[key])
            else:
                value = current.find(S + 'v')
                if value is not None:
                    current.remove(value)
                inline = current.find(S + 'is')
                if inline is not None:
                    current.remove(inline)
                current.append(copy.deepcopy(original_string))
                current.set('t', 'inlineStr')
            repaired_strings.append(ref)
        if repaired_styles or repaired_strings:
            modified_sheets[base.sheets[name]['part']] = target
            changes[base.sheets[name]['part']] = serialized(target)
        input_reports.append({'Sheet': name, 'StylesRestored': repaired_styles,
                              'UnchangedTextCellsCompared': unchanged_text,
                              'RichTextOrSharedStringRestored': repaired_strings})
    changes['xl/styles.xml'] = serialized(target_styles)
    if len(target_strings) != len(X.fromstring(base.parts['xl/sharedStrings.xml'])) or any(n['RichTextOrSharedStringRestored'] for n in input_reports):
        references = 0
        for name, sheet in base.sheets.items():
            xml = modified_sheets.get(sheet['part'])
            if xml is None:
                xml = X.fromstring(base.parts[sheet['part']])
            references += len(xml.findall('s:sheetData/s:row/s:c[@t="s"]', NS))
        target_strings.set('count', str(references))
        target_strings.set('uniqueCount', str(len(target_strings)))
        changes['xl/sharedStrings.xml'] = serialized(target_strings)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(base_path) as source, zipfile.ZipFile(output_path, 'w') as target:
        target.comment = source.comment
        for entry in source.infolist():
            target.writestr(entry, changes.get(entry.filename, source.read(entry.filename)))
    after = VERIFY.Book(output_path)
    rich_errors = []
    after_strings = X.fromstring(after.parts['xl/sharedStrings.xml'])
    for name in INPUTS:
        if name not in old.sheets:
            continue
        for ref, original in old.sheets[name]['rawcells'].items():
            if old.sheets[name]['cells'].get(ref) != after.sheets[name]['cells'].get(ref):
                continue
            original_string = shared_item(source_strings, original)
            current = after.sheets[name]['rawcells'].get(ref)
            current_string = shared_item(after_strings, current) if current is not None else None
            if original_string is not None and (current_string is None or signature(original_string) != signature(current_string)):
                rich_errors.append({'Sheet': name, 'Cell': ref})
    part_diff = [p for p in base.parts if base.parts[p] != after.parts[p]]
    allowed_parts = {'xl/styles.xml', 'xl/sharedStrings.xml'} | {base.sheets[n]['part'] for n in INPUTS if n in base.sheets}
    forbidden = [p for p in part_diff if p not in allowed_parts]
    check = VERIFY.verify(before_path, output_path)
    data_check = VERIFY.verify(base_path, output_path)
    data_errors = [e for e in data_check['Errors'] if e['Kind'] != 'NormalizedCellStyle']
    source_unchanged = VERIFY.digest(before_path.read_bytes()) == original_source_hash
    base_unchanged = VERIFY.digest(base_path.read_bytes()) == original_base_hash
    report = {'Status': 'PASS' if not (check['Errors'] or rich_errors or forbidden or data_errors) and source_unchanged and base_unchanged else 'REVIEW_REQUIRED',
              'Source': str(before_path), 'Base': str(base_path), 'Output': str(output_path),
              'SourceSHA256': VERIFY.digest(before_path.read_bytes()), 'BaseSHA256': VERIFY.digest(base_path.read_bytes()),
              'OutputSHA256': VERIFY.digest(output_path.read_bytes()),
              'ExcelStarted': False, 'SourceUnchanged': source_unchanged, 'BaseUnchanged': base_unchanged,
              'Inputs': input_reports, 'StyleDependenciesAdded': styles.additions,
              'CatalogTextChangesPreserved': text_changes, 'RichTextErrors': rich_errors,
              'ChangedZIPParts': part_diff, 'ForbiddenZIPChanges': forbidden,
              'PartCount': len(after.parts), 'UnchangedZIPPartCount': len(base.parts) - len(part_diff),
              'VBAUnchangedFromBase': base.parts['xl/vbaProject.bin'] == after.parts['xl/vbaProject.bin'],
              'ReadyResultsSheetUnchangedFromBase': base.parts[base.sheets['Results']['part']] == after.parts[after.sheets['Results']['part']],
              'HelpSheetUnchangedFromBase': base.parts[base.sheets['Справка']['part']] == after.parts[after.sheets['Справка']['part']],
              'PreservationErrorCount': len(check['Errors']), 'PreservationErrors': check['Errors'],
              'DataOrDimensionsErrorCountFromBase': len(data_errors), 'DataOrDimensionsErrorsFromBase': data_errors,
              'Limitations': ['Private package only; native Excel reopen and current VBA epoch insertion are coordinated by root.',
                              'Unchanged package parts are byte-identical; modified input sheet XML retains namespaces but is serialized by lxml.']}
    Path(report_path).write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    Path(report_path).with_name('OpenXmlPreservationCompile08ToStyleFixed03.json').write_text(json.dumps(check, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--before', type=Path, required=True)
    parser.add_argument('--base', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    args = parser.parse_args()
    report = restore(args.before, args.base, args.output, args.report)
    print(json.dumps({k: report[k] for k in ('Status', 'OutputSHA256', 'StyleDependenciesAdded', 'ChangedZIPParts',
                                            'PreservationErrorCount', 'DataOrDimensionsErrorCountFromBase',
                                            'RichTextErrors')}, ensure_ascii=True))
    raise SystemExit(report['Status'] != 'PASS')
