"""Переносит проверенную справку из нативной копии без пересохранения данных и VBA.

Входные листы получают только новые адреса прежних ссылок на справку. Стили
добавляются к исходному каталогу, поэтому существующие индексы не меняются.
Рисунки формул сохраняют исходные изображения и получают новые позиции.
"""
import argparse
import copy
import hashlib
import json
import posixpath
import re
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E

M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
X = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'m': M, 'x': X, 'a': A}


def xml(data):
    """Читает XML без нормализации пробелов в пользовательском тексте."""
    return E.fromstring(data)


def dump(root):
    """Сохраняет Unicode и самостоятельное пространство имён каждой части."""
    return E.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def signature(node):
    """Сравнивает содержимое XML без зависимости от имён префиксов и порядка атрибутов."""
    return node.tag, tuple(sorted(node.attrib.items())), node.text, tuple(signature(c) for c in node)


def merge_styles(original, native):
    """Добавляет зависимости нативных стилей, сохраняя все исходные записи на месте."""
    target, source = xml(original), xml(native)
    maps = {}
    old_formats = target.find('m:numFmts', NS)
    source_formats = source.find('m:numFmts', NS)
    formats = {int(n.get('numFmtId')): int(n.get('numFmtId')) for n in old_formats}
    next_format = max([163] + list(formats)) + 1
    for node in source_formats:
        found = next((n for n in old_formats if n.get('formatCode') == node.get('formatCode')), None)
        if found is None:
            found = copy.deepcopy(node)
            found.set('numFmtId', str(next_format))
            next_format += 1
            old_formats.append(found)
        formats[int(node.get('numFmtId'))] = int(found.get('numFmtId'))
    old_formats.set('count', str(len(old_formats)))
    for tag in ('fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs'):
        dest, src = target.find('m:' + tag, NS), source.find('m:' + tag, NS)
        indices = {signature(n): i for i, n in enumerate(dest)}
        mapping = {}
        for index, node in enumerate(src):
            node = copy.deepcopy(node)
            if tag in ('cellStyleXfs', 'cellXfs'):
                for attribute, dependency in (('fontId', 'fonts'), ('fillId', 'fills'), ('borderId', 'borders'), ('xfId', 'cellStyleXfs')):
                    if node.get(attribute) is not None:
                        node.set(attribute, str(maps[dependency][int(node.get(attribute))]))
                number = int(node.get('numFmtId', '0'))
                if number >= 164:
                    node.set('numFmtId', str(formats[number]))
            key = signature(node)
            if key not in indices:
                indices[key] = len(dest)
                dest.append(node)
            mapping[index] = indices[key]
        dest.set('count', str(len(dest)))
        maps[tag] = mapping
    return dump(target), maps['cellXfs']


def help_link(value):
    """Распознаёт только внутреннюю ссылку на лист справки, а не произвольный адрес."""
    return re.fullmatch(r"'?Справка'?!\$?[A-Z]+\$?\d+", value or '') is not None


def drawing_part(parts):
    """Находит рисунки справки через её relationship, без фиксированного номера drawing."""
    rels = xml(parts['xl/worksheets/_rels/sheet3.xml.rels'])
    return next(posixpath.normpath(posixpath.join('xl/worksheets', n.get('Target')))
                for n in rels if n.get('Type').endswith('/drawing'))


def package_help(source, native, target, receipt):
    """Публикует только справку; любая потеря ссылки, рисунка или исходной части прерывает работу."""
    with ZipFile(source) as archive:
        infos = archive.infolist()
        before = {i.filename: archive.read(i) for i in infos}
    with ZipFile(native) as archive:
        candidate = {n: archive.read(n) for n in archive.namelist()}
    parts = dict(before)
    parts['xl/styles.xml'], styles = merge_styles(before['xl/styles.xml'], candidate['xl/styles.xml'])
    help_xml = xml(candidate['xl/worksheets/sheet3.xml'])
    strings = list(xml(candidate['xl/sharedStrings.xml']))
    for cell in help_xml.findall('.//m:c', NS):
        if cell.get('s') is not None:
            cell.set('s', str(styles[int(cell.get('s'))]))
        if cell.get('t') == 's':
            index = int(cell.find('m:v', NS).text)
            for child in list(cell):
                cell.remove(child)
            cell.set('t', 'inlineStr')
            text = E.SubElement(cell, '{' + M + '}is')
            for child in strings[index]:
                text.append(copy.deepcopy(child))
    for node in help_xml.findall('m:cols/m:col', NS) + help_xml.findall('m:sheetData/m:row', NS):
        if node.get('s') is not None:
            node.set('s', str(styles[int(node.get('s'))]))
        if node.get('style') is not None:
            node.set('style', str(styles[int(node.get('style'))]))
    original_help = xml(before['xl/worksheets/sheet3.xml'])
    help_xml.find('m:drawing', NS).set('{' + R + '}id', original_help.find('m:drawing', NS).get('{' + R + '}id'))
    parts['xl/worksheets/sheet3.xml'] = dump(help_xml)
    link_count = 0
    for name in before:
        if name == 'xl/worksheets/sheet3.xml' or not re.fullmatch(r'xl/worksheets/sheet\d+\.xml', name):
            continue
        old, new = xml(before[name]), xml(candidate[name])
        old_links = {n.get('ref'): n for n in old.findall('m:hyperlinks/m:hyperlink', NS) if help_link(n.get('location'))}
        new_links = {n.get('ref'): n for n in new.findall('m:hyperlinks/m:hyperlink', NS) if help_link(n.get('location'))}
        assert old_links.keys() == new_links.keys(), ('Changed input link anchors', name)
        for ref, link in old_links.items():
            link.set('location', new_links[ref].get('location'))
            link_count += 1
        if old_links:
            parts[name] = dump(old)
    old, new = xml(before['xl/workbook.xml']), xml(candidate['xl/workbook.xml'])
    source_name = new.find("m:definedNames/m:definedName[@name='rngSectionRecipeGuide']", NS)
    old.find("m:definedNames/m:definedName[@name='rngSectionRecipeGuide']", NS).text = source_name.text
    parts['xl/workbook.xml'] = dump(old)
    old_drawing_name, new_drawing_name = drawing_part(before), drawing_part(candidate)
    old_drawing, new_drawing = xml(before[old_drawing_name]), xml(candidate[new_drawing_name])
    anchors = {n.find('x:pic/x:nvPicPr/x:cNvPr', NS).get('name'): n for n in new_drawing}
    names = []
    for anchor in old_drawing:
        picture = anchor.find('x:pic', NS)
        name = picture.find('x:nvPicPr/x:cNvPr', NS).get('name')
        native_anchor = anchors[name]
        assert anchor.tag == native_anchor.tag
        for marker in ('from', 'to'):
            previous = anchor.find('x:' + marker, NS)
            anchor.replace(previous, copy.deepcopy(native_anchor.find('x:' + marker, NS)))
        transform = picture.find('x:spPr/a:xfrm', NS)
        native_transform = native_anchor.find('x:pic/x:spPr/a:xfrm', NS)
        transform.getparent().replace(transform, copy.deepcopy(native_transform))
        assert int(native_transform.find('a:ext', NS).get('cy')) > 0, name
        names.append(name)
    assert set(names) == anchors.keys() and len(names) == 76
    parts[old_drawing_name] = dump(old_drawing)
    changed = [n for n in before if before[n] != parts[n]]
    allowed = {'xl/styles.xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml', old_drawing_name}
    assert set(changed) <= allowed, changed
    assert before['xl/vbaProject.bin'] == parts['xl/vbaProject.bin']
    assert not Path(target).exists()
    with ZipFile(target, 'w') as archive:
        for info in infos:
            archive.writestr(info, parts[info.filename])
    result = {'Status': 'PASS', 'SourceSHA256': hashlib.sha256(Path(source).read_bytes()).hexdigest().upper(),
              'TargetSHA256': hashlib.sha256(Path(target).read_bytes()).hexdigest().upper(),
              'ChangedParts': changed, 'HelpLinks': link_count, 'Pictures': len(names),
              'VBAByteIdentical': True, 'MediaByteIdentical': True, 'OtherPartsByteIdentical': True}
    Path(receipt).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(result))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ('source', 'native', 'target', 'receipt'):
        parser.add_argument(name)
    arguments = parser.parse_args()
    package_help(arguments.source, arguments.native, arguments.target, arguments.receipt)
