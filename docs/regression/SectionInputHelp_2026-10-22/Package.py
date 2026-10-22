"""Publish only the edited help cells and compiled workbook event, preserving native XLSM parts."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile
import argparse, hashlib, json, re, sys
from lxml import etree as E

sys.stdout.reconfigure(encoding='utf-8')
HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
S = '{' + NS['s'] + '}'
OLD_TITLE = 'Справка по настройкам листа Config'
NEW_TITLE = 'Справка по настройкам Config и листа «Сечения»'
native_dir = HERE / 'Native02'
edit = json.loads((native_dir / 'NativeEdit.json').read_text(encoding='utf-8-sig'))
catalog = json.loads((native_dir / 'Catalog.json').read_text(encoding='utf-8-sig'))
definitions = {item['Key']: item for item in catalog}
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest().upper()


def read_book(path):
    with ZipFile(path) as z:
        parts = {i.filename: z.read(i) for i in z.infolist()}
    wb = E.fromstring(parts['xl/workbook.xml'])
    rels = {r.get('Id'): r.get('Target') for r in E.fromstring(parts['xl/_rels/workbook.xml.rels'])}
    sheets = {s.get('name'): 'xl/' + rels[s.get('{' + NS['r'] + '}id')]
              for s in wb.find('s:sheets', NS)}
    strings = [''.join(x.itertext()) for x in E.fromstring(parts['xl/sharedStrings.xml'])]
    return parts, sheets, strings, wb


def values(doc, strings):
    result = {}
    for c in doc.findall('s:sheetData/s:row/s:c', NS):
        if c.get('t') == 'inlineStr':
            result[c.get('r')] = ''.join(c.find('s:is', NS).itertext())
        else:
            v = c.find('s:v', NS)
            result[c.get('r')] = strings[int(v.text)] if c.get('t') == 's' else (v.text if v is not None else '')
    return result


def write_text(cell, text):
    for child in list(cell):
        cell.remove(child)
    cell.set('t', 'inlineStr')
    E.SubElement(E.SubElement(cell, S + 'is'), S + 't').text = text


parser = argparse.ArgumentParser()
parser.add_argument('--verify', type=Path)
args = parser.parse_args()
baseline = HERE / 'Baseline.xlsm'
assert sha(baseline) == edit['SourceSHA256']
old, sheets, strings, wb = read_book(baseline)
help_part = sheets['Справка']
original_help = E.fromstring(old[help_part])
before_values = values(original_help, strings)
new_help = deepcopy(original_help)
cells = {c.get('r'): c for c in new_help.findall('s:sheetData/s:row/s:c', NS)}
changes = {c['Address']: c for c in edit['Changes'] if c['Before'] != c['After']}
for address, change in changes.items():
    assert before_values[address] == change['Before'], address
    write_text(cells[address], change['After'])

# Add only a blue underlined variant of the existing paragraph style.
styles = E.fromstring(old['xl/styles.xml'])
fonts = styles.find('s:fonts', NS)
xfs = styles.find('s:cellXfs', NS)
link_styles = {}
links = new_help.find('s:hyperlinks', NS)
assert links is not None, 'Expected existing help navigation'
new_links = {}
for block in edit['Blocks']:
    for row, height in block['Heights'].items():
        target = new_help.find(f's:sheetData/s:row[@r="{row}"]', NS)
        target.set('ht', format(height, '.12g'))
        target.set('customHeight', '1')
    if not block['RelatedKey']:
        continue
    address = 'B' + str(block['Last'])
    style_id = int(cells[address].get('s', '0'))
    if style_id not in link_styles:
        xf = deepcopy(xfs[style_id])
        font = deepcopy(fonts[int(xf.get('fontId', '0'))])
        for child in list(font):
            if child.tag in (S + 'color', S + 'u'):
                font.remove(child)
        E.SubElement(font, S + 'color', rgb='FF0000FF')
        E.SubElement(font, S + 'u', val='single')
        xf.set('fontId', str(len(fonts)))
        fonts.append(font)
        link_styles[style_id] = len(xfs)
        xfs.append(xf)
    cells[address].set('s', str(link_styles[style_id]))
    for link in list(links):
        if link.get('ref') == address:
            links.remove(link)
    common = next(b for b in edit['Blocks'] if b['Key'] == block['RelatedKey'])
    location = f"'Справка'!A{common['First']}"
    E.SubElement(links, S + 'hyperlink', ref=address, location=location,
                 display=definitions[block['Key']]['Lines'][-1])
    new_links[address] = location
fonts.set('count', str(len(fonts)))
xfs.set('count', str(len(xfs)))
encode = lambda doc: E.tostring(doc, encoding='UTF-8', xml_declaration=True, standalone=True)
new_parts = dict(old)
new_parts[help_part] = encode(new_help)
new_parts['xl/styles.xml'] = encode(styles)
native_binary = native_dir / 'vbaProject.bin'
if native_binary.exists():
    new_parts['xl/vbaProject.bin'] = native_binary.read_bytes()
else:
    native_book = native_dir / 'NativeCandidate.xlsm'
    if not native_book.exists():
        native_book = args.verify or ROOT / 'workbook/output/RC_Section_NDM.xlsm'
    with ZipFile(native_book) as z:
        new_parts['xl/vbaProject.bin'] = z.read('xl/vbaProject.bin')
assert hashlib.sha256(new_parts['xl/vbaProject.bin']).hexdigest().upper() == edit['CompiledVbaSHA256']
candidate = args.verify or HERE / 'Release/RC_Section_NDM.xlsm'
if not args.verify:
    candidate.parent.mkdir(exist_ok=True)
    with ZipFile(baseline) as src, ZipFile(candidate, 'w') as dst:
        for info in src.infolist():
            dst.writestr(info, new_parts[info.filename])

actual, actual_sheets, actual_strings, actual_wb = read_book(candidate)
assert actual.keys() == old.keys(), 'Native package inventory changed'
allowed = {help_part, 'xl/styles.xml', 'xl/vbaProject.bin'}
changed_parts = [name for name in old if actual[name] != old[name]]
assert set(changed_parts) == allowed, changed_parts
assert all(actual[p] == new_parts[p] for p in allowed), 'Unexpected published edits'
old_styles = E.fromstring(old['xl/styles.xml'])
actual_styles = E.fromstring(actual['xl/styles.xml'])
for group in old_styles:
    other = actual_styles.find(group.tag)
    assert other is not None
    if group.tag in (S + 'fonts', S + 'cellXfs'):
        assert all(E.tostring(a, method='c14n') == E.tostring(b, method='c14n')
                   for a, b in zip(group, other)), 'An existing style changed'
    else:
        assert E.tostring(group, method='c14n') == E.tostring(other, method='c14n')
doc = E.fromstring(actual[help_part])
after_values = values(doc, actual_strings)
assert before_values.keys() == after_values.keys(), 'Help cell inventory changed'
for address in before_values:
    expected = changes[address]['After'] if address in changes else before_values[address]
    assert after_values[address] == expected, address
assert NEW_TITLE in after_values.values() and OLD_TITLE not in after_values.values()
actual_links = {l.get('ref'): l.get('location') for l in doc.findall('s:hyperlinks/s:hyperlink', NS)}
assert all(actual_links.get(a) == loc for a, loc in new_links.items())
for link in original_help.findall('s:hyperlinks/s:hyperlink', NS):
    assert actual_links[link.get('ref')] == link.get('location'), 'An existing help link moved'
for block in edit['Blocks']:
    item = definitions[block['Key']]
    assert after_values['A' + str(block['First'])] == item['Title']
    for index, text in enumerate(item['Lines'], block['First'] + 1):
        assert after_values['B' + str(index)] == text

# Registry, actual named tables, and actual header destinations must agree.
geometry = json.loads((ROOT / 'tools/build_workbook/GeometryTables.json').read_text(encoding='utf-8-sig'))
registry = (ROOT / 'src/Excel/CSectionTypeRegistry.cls').read_text(encoding='utf-8-sig')
registered = re.findall(r'RegisterRecipe "([^"]+)".*?Array\("([^"]+)"\)', registry)
assert set(registered) == {(d['Type'], d['RangeName']) for d in geometry}
names = {n.get('name'): n.text for n in wb.find('s:definedNames', NS)}
section_doc = E.fromstring(actual[sheets['Сечения']])
section_links = {l.get('ref'): l for l in section_doc.findall('s:hyperlinks/s:hyperlink', NS)}
section_values = values(section_doc, strings)
coverage = []
for definition in geometry:
    match = re.fullmatch(r"'?Сечения'?!\$([A-Z]+)\$(\d+):\$[A-Z]+\$\d+", names[definition['RangeName']])
    assert match, definition['RangeName']
    address = match[1] + str(int(match[2]) - 1)
    link = section_links[address]
    block = next(b for b in edit['Blocks'] if definitions[b['Key']]['Title'] == definition['Title'])
    assert link.get('location') == f"'Справка'!A{block['First']}"
    assert '(Подробнее)' in section_values[address]
    coverage.append({'Type': definition['Type'], 'Range': definition['RangeName'],
                     'Header': address, 'Target': link.get('location'),
                     'CommonHelpLink': 'B' + str(block['Last'])})
paragraphs = [t for item in catalog for t in item['Lines'] if not t.startswith('Подробнее:')]
assert len(paragraphs) == len(set(paragraphs)), 'Repeated descriptive paragraphs'
obsolete = r'Tapered|PointedRect(?!angle)|RoundedRectangle|RectSet|rngRebarInput|общ[а-я]+ таблиц[а-я]+ арм'
assert not any(re.search(obsolete, t, re.I) for t in after_values.values() if isinstance(t, str))
before_chars = sum(len(c['Before']) for c in edit['Changes'])
after_chars = sum(len(c['After']) for c in edit['Changes'])
result = {'Passed': True, 'WorkbookSHA256': sha(candidate), 'BaselineSHA256': sha(baseline),
          'ChangedParts': changed_parts, 'UnchangedPackageParts': len(old) - len(changed_parts),
          'AllUserInputAndResultsPartsByteIdentical': True,
          'UnchangedDrawingParts': sum(p.startswith(('xl/drawings/', 'xl/charts/', 'xl/media/')) for p in old),
          'ChangedHelpCells': len(changes), 'RegisteredTypes': len(registered), 'Coverage': coverage,
          'CommonLinks': new_links, 'DescriptiveParagraphs': len(paragraphs),
          'ExactDuplicateDescriptions': 0, 'ObsoleteSectionNamesAndSharedRebarReferences': 0,
          'BeforeCharacters': before_chars, 'AfterCharacters': after_chars}
(HERE / ('PublishedVerification.json' if args.verify else 'Acceptance.json')).write_text(
    json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps(result, ensure_ascii=False, indent=2))
