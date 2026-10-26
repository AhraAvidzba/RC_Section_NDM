"""ДЛЯ ТЕСТОВ: read-only декомпозиция стилей после native обновления hyperlinks."""
from pathlib import Path
from xml.etree import ElementTree as E
import collections
import importlib.util
import json

ROOT = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location('check', ROOT / 'TestVerifyOpenXmlPreservation.py')
CHECK = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(CHECK)
NS, S = CHECK.NS, '{' + CHECK.NS['s'] + '}'


def style(book, index):
    root = E.fromstring(book.parts['xl/styles.xml'])
    xf = list(root.find('s:cellXfs', NS))[index]
    bases = list(root.find('s:cellStyleXfs', NS))
    formats = {int(x.get('numFmtId')): x.get('formatCode') for x in root.findall('s:numFmts/s:numFmt', NS)}

    def item(node, parent):
        result = {'Attributes': {k: v for k, v in node.attrib.items() if k not in ('fontId', 'fillId', 'borderId', 'numFmtId', 'xfId')},
                  'Children': CHECK.canon(node)[3]}
        for attribute, group in [('fontId', 'fonts'), ('fillId', 'fills'), ('borderId', 'borders')]:
            result[group] = CHECK.canon(list(root.find(S + group))[int(node.get(attribute, '0'))])
        result['NumberFormat'] = formats.get(int(node.get('numFmtId', '0')), 'BUILTIN:' + node.get('numFmtId', '0'))
        if parent:
            result['Base'] = item(bases[int(node.get('xfId', '0'))], False)
        return result

    return {'Index': index, 'Raw': E.tostring(xf, encoding='unicode'), 'Resolved': item(xf, True)}


def shared_string(book, cell):
    if cell.get('t') == 's':
        return book.string_nodes[int(cell.find(S + 'v').text)]
    if cell.get('t') == 'inlineStr':
        return cell.find(S + 'is')
    return None


def diagnose(before, after, report_path):
    old, new = CHECK.Book(before), CHECK.Book(after)
    for book in (old, new):
        book.string_nodes = list(E.fromstring(book.parts['xl/sharedStrings.xml']))
    report = CHECK.verify(before, after)
    grouped = collections.defaultdict(list)
    for error in report['Errors']:
        if error['Kind'] != 'NormalizedCellStyle':
            continue
        name, ref = error['Sheet'], error['BeforeCell']
        a = int(old.sheets[name]['rawcells'][ref].get('s', '0'))
        b = int(new.sheets[name]['rawcells'][ref].get('s', '0'))
        grouped[(a, b)].append([name, ref])
    groups = []
    for (a, b), cells in grouped.items():
        sa, sb = style(old, a), style(new, b)
        different = [k for k in sa['Resolved'] if sa['Resolved'][k] != sb['Resolved'][k]]
        groups.append({'SourceStyle': a, 'AfterStyle': b, 'Count': len(cells), 'Cells': cells,
                       'DifferentComponents': different, 'Before': sa, 'After': sb})
    groups.sort(key=lambda group: -group['Count'])
    strings = []
    for name in ('Config', 'Сечения', 'Расчет'):
        for ref, cell in old.sheets[name]['rawcells'].items():
            if ref not in new.sheets[name]['rawcells'] or old.sheets[name]['cells'].get(ref) != new.sheets[name]['cells'].get(ref):
                continue
            current = new.sheets[name]['rawcells'][ref]
            a, b = shared_string(old, cell), shared_string(new, current)
            if (a is not None or b is not None) and CHECK.canon(a) != CHECK.canon(b):
                strings.append({'Sheet': name, 'Cell': ref,
                                'Before': E.tostring(a, encoding='unicode') if a is not None else None,
                                'After': E.tostring(b, encoding='unicode') if b is not None else None})
    proof = {'Status': 'READ_ONLY_DIAGNOSIS', 'SourceSHA256': CHECK.digest(old.path.read_bytes()),
             'AfterSHA256': CHECK.digest(new.path.read_bytes()),
             'ErrorKinds': dict(collections.Counter(error['Kind'] for error in report['Errors'])),
             'StyleCells': sum(group['Count'] for group in groups), 'Groups': groups,
             'UnchangedTextRichStringDifferences': strings, 'ExcelStarted': False, 'BooksEdited': False}
    Path(report_path).write_text(json.dumps(proof, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    return proof


if __name__ == '__main__':
    proof = diagnose(ROOT / 'ReleaseCandidate05/RC_Section_NDM.xlsm', ROOT / 'HyperlinkNative04/RC_Section_NDM.xlsm',
                     ROOT / 'HyperlinkNative04/StyleDiagnosis.json')
    print(json.dumps({'ErrorKinds': proof['ErrorKinds'],
                      'Groups': [{k: group[k] for k in ('SourceStyle', 'AfterStyle', 'Count', 'DifferentComponents')} for group in proof['Groups']],
                      'RichTextDifferences': len(proof['UnchangedTextRichStringDifferences'])}, ensure_ascii=True))
