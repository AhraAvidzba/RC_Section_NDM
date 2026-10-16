"""Переносит проверенный VBA и точечное уточнение Справки без пересохранения данных Excel."""
import hashlib
import importlib.util
import json
import re
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
D = ROOT / 'docs/regression/ExcelGuardGuide_2026-10-13/RecipeSimplification_2026-10-16'
spec = importlib.util.spec_from_file_location('presentation', Path(__file__).with_name('Verify-GuidePresentation.py'))
presentation = importlib.util.module_from_spec(spec)
spec.loader.exec_module(presentation)
reader = presentation.reader


def normalize(text):
    return ''.join(part if part.startswith(('"', "'")) else part.lower() for part in re.findall(
        r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27\n]+|\n', text)).strip()


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def package():
    baseline = presentation.components(D / 'BaselineNativeVBA.txt')
    probe = presentation.components(D / 'BaselineProbeNativeVBA.txt')
    candidate = presentation.components(D / 'CandidateNativeVBA.txt')
    assert len(baseline) == 157 and len(candidate) == 158
    assert candidate.keys() - baseline.keys() == {'modSectionRecipeRows'}
    assert not baseline.keys() - candidate.keys()
    assert baseline.keys() == probe.keys()
    assert [name for name in baseline if normalize(baseline[name][2]) != normalize(probe[name][2])] == ['modTestAudit02Snapshots']
    changed = {name for name in baseline if normalize(baseline[name][2]) != normalize(candidate[name][2])}
    assert changed == {'CSectionRecipeCircle', 'CSectionRecipeRectSet', 'CSectionRecipeRoundedRectangle',
                       'CSectionRecipeHollowRectangle', 'modSectionRecipeApiGuide', 'modTestAudit02Api', 'modTestAudit02Snapshots'}, changed
    equality = json.loads((D / 'CandidateSourceEquality.json').read_text(encoding='utf-8-sig'))
    assert equality['Passed']
    baseline_file, candidate_file = D / 'Baseline.xlsm', D / 'Candidate/RecipeComparison.xlsm'
    old, new = reader.read(baseline_file), reader.read(candidate_file)
    assert old['names'] == new['names']
    layout = json.loads((ROOT / 'docs/geometry_refactor/SECTION_RECIPE_GUIDE_LAYOUT.json').read_text(encoding='utf-8-sig'))
    start, end = map(int, re.fullmatch(r"'?Справка'?!\$A\$(\d+):\$F\$(\d+)", old['names']['rngSectionRecipeGuide']).groups())
    assert end - start + 1 == len(layout['rows'])
    changes = {'A'+str(start+i): row['text'] for i, row in enumerate(layout['rows'])
               if old['sheets']['Справка']['cells'].get('A'+str(start+i), {}).get('value', '') != row['text']}
    assert len(changes) == 1
    height_changes = {}
    with ZipFile(baseline_file) as original, ZipFile(candidate_file) as compiled:
        book = ET.fromstring(original.read('xl/workbook.xml'))
        relations = {item.attrib['Id']: item.attrib['Target'] for item in ET.fromstring(original.read('xl/_rels/workbook.xml.rels'))}
        sheet = next(item for item in book.find(reader.S+'sheets') if item.attrib['name'] == 'Справка')
        target = relations[sheet.attrib[reader.R+'id']]
        part = target.lstrip('/') if target.startswith('/') else 'xl/'+target
        xml = original.read(part).decode('utf-8')
        for address, text in changes.items():
            assert new['sheets']['Справка']['cells'][address]['value'] == text
            pattern = r'<c\b[^>]*\br="'+re.escape(address)+r'"[^>]*(?:/>|>.*?</c>)'
            matches = list(re.finditer(pattern, xml, re.S))
            assert len(matches) == 1, address
            opening = matches[0].group().split('>', 1)[0].rstrip('/')
            opening = re.sub(r'\s+t="[^"]*"', '', opening)+' t="inlineStr">'
            assert old['sheets']['Справка']['cells'][address]['formula'] is None
            replacement = opening+'<is><t>'+escape(text)+'</t></is></c>'
            xml = xml[:matches[0].start()]+replacement+xml[matches[0].end():]
            row_number = int(re.search(r'\d+$', address).group())
            old_height = old['sheets']['Справка']['rows'][row_number].get('ht')
            new_height = new['sheets']['Справка']['rows'][row_number].get('ht')
            if old_height != new_height:
                assert float(new_height) >= float(old_height)
                row_pattern = r'<row\b[^>]*\br="'+str(row_number)+r'"[^>]*>'
                row_matches = list(re.finditer(row_pattern, xml))
                assert len(row_matches) == 1
                opening = row_matches[0].group()
                assert len(re.findall(r'\bht="[^"]*"', opening)) == 1
                replacement = re.sub(r'\bht="[^"]*"', 'ht="'+new_height+'"', opening)
                xml = xml[:row_matches[0].start()]+replacement+xml[row_matches[0].end():]
                height_changes[str(row_number)] = {'Before': old_height, 'After': new_height}
        accepted = D / 'Accepted.xlsm'
        assert not accepted.exists(), 'Do not overwrite an existing accepted package.'
        with ZipFile(accepted, 'w') as output:
            for entry in original.infolist():
                data = compiled.read(entry.filename) if entry.filename == 'xl/vbaProject.bin' else original.read(entry.filename)
                if entry.filename == part:
                    data = xml.encode('utf-8')
                output.writestr(entry, data)
    accepted_state = reader.read(accepted)
    checked = 0
    for sheet, old_sheet in old['sheets'].items():
        actual = accepted_state['sheets'][sheet]
        expected_rows = {number: dict(attributes) for number, attributes in old_sheet['rows'].items()}
        if sheet == 'Справка':
            for row_number, heights in height_changes.items():
                expected_rows[int(row_number)]['ht'] = heights['After']
        assert expected_rows == actual['rows'], 'Unexpected row height or formatting change: '+sheet
        assert old_sheet['cells'].keys() == actual['cells'].keys()
        for address, cell in old_sheet['cells'].items():
            expected = dict(cell)
            if sheet == 'Справка' and address in changes:
                expected['value'] = changes[address]
            assert expected == actual['cells'][address], sheet+'!'+address
            checked += 1
    assert old['names'] == accepted_state['names'] and old['drawings'] == accepted_state['drawings']
    with ZipFile(baseline_file) as a, ZipFile(accepted) as b:
        assert a.namelist() == b.namelist()
        changed_parts = [name for name in a.namelist() if a.read(name) != b.read(name)]
        assert set(changed_parts) == {'xl/vbaProject.bin', part}
    assert not presentation.embedded_markers(accepted, D / 'CandidateNativeVBA.txt')
    receipt = {'Status': 'RECIPE_PACKAGE_PASS', 'BaselineSHA256': sha(baseline_file),
               'AcceptedSHA256': sha(accepted), 'CandidateSHA256': sha(candidate_file),
               'NativeSHA256': sha(D / 'CandidateNativeVBA.txt'), 'NativeComponents': len(candidate),
               'ChangedComponents': sorted(changed), 'AddedComponents': ['modSectionRecipeRows'],
               'BaselineProbeChanges': ['modTestAudit02Snapshots'], 'ChangedZIPParts': changed_parts,
               'ChangedHelpCells': changes, 'CellsChecked': checked, 'HelpRowHeightChanges': height_changes,
               'AllOtherRowHeightsUnchanged': True,
               'ConfigAndAllOtherUserDataPreserved': True}
    (D / 'Package.json').write_text(json.dumps(receipt, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps(receipt, ensure_ascii=False))


if __name__ == '__main__':
    package()
