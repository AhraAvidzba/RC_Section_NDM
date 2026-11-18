"""Обновляет ссылки источников в закрытой XLSM без изменения VBA и данных.

Источники без публичного URL в общем каталоге остаются обычными ссылками
на литературу: их названия и страницы сохраняются, гиперссылки удаляются.
"""
import argparse
import copy
import hashlib
import json
import re
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
M = '{' + NS['m'] + '}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def sha(data):
    """Возвращает контрольную сумму исходных или изменённых байтов."""
    return hashlib.sha256(data).hexdigest().upper()


def update(source, target, catalog_path):
    """Меняет только подписи ссылок, их шрифты и зависимую цепочку пересчёта."""
    if source.resolve() == target.resolve() or target.exists():
        raise ValueError('Нужен отдельный новый выходной файл.')
    catalog = json.loads(catalog_path.read_text(encoding='utf-8-sig'))
    disabled = [x['caption'] for x in catalog if x.get('id') and not x.get('url')]
    with ZipFile(source) as archive:
        infos = archive.infolist()
        parts = {x.filename: archive.read(x) for x in infos}
        comment = archive.comment
    original = dict(parts)
    workbook = E.fromstring(parts['xl/workbook.xml'])
    relationships = E.fromstring(parts['xl/_rels/workbook.xml.rels'])
    targets = {x.get('Id'): x.get('Target') for x in relationships}
    sheet = next(x for x in workbook.find('m:sheets', NS) if x.get('name') == 'Справка')
    part = targets[sheet.get(REL + 'id')]
    part = part.lstrip('/') if part.startswith('/') else 'xl/' + part
    guide = E.fromstring(parts[part])
    styles = E.fromstring(parts['xl/styles.xml'])
    fonts = styles.find('m:fonts', NS)
    formats = styles.find('m:cellXfs', NS)
    style_map, changes = {}, []
    formula_pattern = re.compile(r'^HYPERLINK\("((?:[^"]|"")*)","((?:[^"]|"")*)"\)$')
    for cell in guide.findall('.//m:sheetData/m:row/m:c', NS):
        formula = cell.find('m:f', NS)
        if formula is None:
            continue
        match = formula_pattern.fullmatch(formula.text or '')
        if not match:
            continue
        address, caption = (x.replace('""', '"') for x in match.groups())
        if not any(caption == name or caption.startswith('Источник: ' + name + '.')
                   or caption.startswith('Источник: ' + name + ' —') for name in disabled):
            continue
        text = caption.removesuffix(' — открыть документ')
        old_style = int(cell.get('s', '0'))
        if old_style not in style_map:
            font = copy.deepcopy(fonts[int(formats[old_style].get('fontId', '0'))])
            for node in list(font):
                if node.tag in (M + 'u', M + 'color'):
                    font.remove(node)
            E.SubElement(font, M + 'color', rgb='FF000000')
            font_xml = E.tostring(font)
            font_id = next((i for i, x in enumerate(fonts) if E.tostring(x) == font_xml), len(fonts))
            if font_id == len(fonts):
                fonts.append(font)
            formatting = copy.deepcopy(formats[old_style])
            formatting.set('fontId', str(font_id))
            format_xml = E.tostring(formatting)
            style_id = next((i for i, x in enumerate(formats) if E.tostring(x) == format_xml), len(formats))
            if style_id == len(formats):
                formats.append(formatting)
            style_map[old_style] = style_id
        cell.set('s', str(style_map[old_style]))
        for node in list(cell):
            if node.tag in (M + 'f', M + 'v', M + 'is'):
                cell.remove(node)
        cell.set('t', 'inlineStr')
        E.SubElement(E.SubElement(cell, M + 'is'), M + 't').text = text
        changes.append({'Cell': cell.get('r'), 'RemovedURL': address, 'Text': text})
    if changes:
        parts[part] = E.tostring(guide, encoding='UTF-8', xml_declaration=True, standalone=True)
        fonts.set('count', str(len(fonts)))
        formats.set('count', str(len(formats)))
        parts['xl/styles.xml'] = E.tostring(styles, encoding='UTF-8', xml_declaration=True, standalone=True)
        if 'xl/calcChain.xml' in parts:
            chain = E.fromstring(parts['xl/calcChain.xml'])
            removed = {x['Cell'] for x in changes}
            current_sheet, retained_sheet = None, None
            for node in list(chain):
                if node.get('i') is not None:
                    current_sheet = node.get('i')
                if current_sheet == sheet.get('sheetId') and node.get('r') in removed:
                    chain.remove(node)
                    continue
                if current_sheet != retained_sheet and current_sheet is not None:
                    node.set('i', current_sheet)
                retained_sheet = current_sheet
            parts['xl/calcChain.xml'] = E.tostring(chain, encoding='UTF-8', xml_declaration=True, standalone=True)
    changed_parts = [name for name in parts if parts[name] != original[name]]
    assert set(changed_parts) <= {part, 'xl/styles.xml', 'xl/calcChain.xml'}
    assert parts['xl/vbaProject.bin'] == original['xl/vbaProject.bin']
    with ZipFile(target, 'w') as archive:
        archive.comment = comment
        for info in infos:
            archive.writestr(info, parts[info.filename])
    return {'Status': 'SOURCE_LINKS_UPDATED', 'OriginalSHA256': sha(source.read_bytes()),
            'OutputSHA256': sha(target.read_bytes()), 'DisabledSources': disabled,
            'RemovedHyperlinks': changes, 'ChangedParts': changed_parts,
            'VBAAndOtherPartsByteIdentical': True}


def main():
    """Принимает закрытую книгу, отдельный результат и путь отчёта обновления."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('target', type=Path)
    parser.add_argument('--report', required=True, type=Path)
    parser.add_argument('--catalog', type=Path, default=Path(__file__).with_name('NormativeHelpSources.json'))
    args = parser.parse_args()
    result = update(args.source, args.target, args.catalog)
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'Status': result['Status'], 'Removed': len(result['RemovedHyperlinks']),
                      'ChangedParts': result['ChangedParts']}, ensure_ascii=False))


if __name__ == '__main__':
    main()
