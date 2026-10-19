"""Упаковывает скомпилированный VBA и один абзац Справки, сохраняя остальные части исходной книги."""
import hashlib
import importlib.util
import json
import re
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
D = ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/SegmentNaming_2026-10-16'
spec = importlib.util.spec_from_file_location('presentation', Path(__file__).with_name('Verify-GuidePresentation.py'))
presentation = importlib.util.module_from_spec(spec)
spec.loader.exec_module(presentation)
reader = presentation.reader


def normalize(body):
    return ''.join(part if part.startswith(('"', "'")) else part.lower() for part in re.findall(
        r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27\n]+|\n', body)).strip()


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def package():
    before = presentation.components(D/'BaselineNativeVBA.txt')
    candidate = presentation.components(D/'CandidateNativeVBA.txt')
    tested = presentation.components(D/'TestedNativeVBA.txt')
    probe = presentation.components(D/'BaselineProbeNativeVBA.txt')
    live = presentation.components(D/'LiveFinalNativeVBA.txt')
    assert before.keys() == candidate.keys() == tested.keys() == probe.keys() and len(before) == 158
    assert all(before[name][0] == candidate[name][0] for name in before)
    changed = {name for name in before if normalize(before[name][2]) != normalize(candidate[name][2])}
    expected = {'CGeometryPath', 'CPolygonSpec', 'CSectionDefinitionBuilder', 'CSectionGeometryQuery',
                'CSectionRecipePointedRect', 'CSectionRecipeHollowRectangle', 'modSectionRecipeApiGuide',
                'modTestPolygonSpec', 'modTestAudit02Api', 'modTestAudit02Snapshots'}
    assert changed == expected, changed
    assert all(normalize(candidate[name][2]) == normalize(tested[name][2]) for name in candidate), 'Tested code differs'
    assert {name for name in before if normalize(before[name][2]) != normalize(probe[name][2])} == {'modTestAudit02Snapshots'}
    assert live.keys() == before.keys() and all(normalize(live[name][2]) == normalize(before[name][2]) for name in before), 'User VBA changed'
    assert json.loads((D/'Acceptance.json').read_text(encoding='utf-8-sig'))['Status'] == 'SEGMENT_NAMING_PASS'
    assert json.loads((D/'CandidateSourceEquality.json').read_text(encoding='utf-8-sig'))['Passed']
    original, compiled = D/'LiveFinalBeforeClose.xlsm', D/'Candidate/SegmentNaming.xlsm'
    old, new = reader.read(original), reader.read(compiled)
    layout = json.loads((ROOT/'docs/geometry_refactor/SECTION_RECIPE_GUIDE_LAYOUT.json').read_text(encoding='utf-8'))
    first, last = map(int, re.fullmatch(r"'?Справка'?!\$A\$(\d+):\$F\$(\d+)", old['names']['rngSectionRecipeGuide']).groups())
    assert last-first+1 == len(layout['rows'])
    changes = {'A'+str(first+i): row['text'] for i,row in enumerate(layout['rows'])
               if (old['sheets']['Справка']['cells'].get('A'+str(first+i),{}).get('value') or '') != row['text']}
    assert len(changes) == 1, changes
    height_changes = {}
    accepted = D/'Accepted.xlsm'
    assert not accepted.exists(), 'Keep the existing accepted artifact'
    with ZipFile(original) as baseline, ZipFile(compiled) as native:
        book = ET.fromstring(baseline.read('xl/workbook.xml'))
        relations = {item.attrib['Id']:item.attrib['Target'] for item in ET.fromstring(baseline.read('xl/_rels/workbook.xml.rels'))}
        sheet = next(item for item in book.find(reader.S+'sheets') if item.attrib['name'] == 'Справка')
        target = relations[sheet.attrib[reader.R+'id']]
        part = target.lstrip('/') if target.startswith('/') else 'xl/'+target
        xml = baseline.read(part).decode('utf-8')
        for address,text in changes.items():
            assert new['sheets']['Справка']['cells'][address]['value'] == text
            matches = list(re.finditer(r'<c\b[^>]*\br="'+re.escape(address)+r'"[^>]*(?:/>|>.*?</c>)', xml, re.S))
            assert len(matches) == 1
            opening = re.sub(r'\s+t="[^"]*"', '', matches[0].group().split('>',1)[0].rstrip('/'))+' t="inlineStr">'
            replacement = opening+'<is><t>'+escape(text)+'</t></is></c>'
            xml = xml[:matches[0].start()]+replacement+xml[matches[0].end():]
            row = int(re.search(r'\d+$',address).group())
            old_height = old['sheets']['Справка']['rows'][row].get('ht')
            new_height = new['sheets']['Справка']['rows'][row].get('ht')
            assert 0 < float(new_height) <= 60
            if old_height != new_height:
                matches = list(re.finditer(r'<row\b[^>]*\br="'+str(row)+r'"[^>]*>', xml))
                assert len(matches) == 1 and len(re.findall(r'\bht="[^"]*"',matches[0].group())) == 1
                replacement = re.sub(r'\bht="[^"]*"','ht="'+new_height+'"',matches[0].group())
                xml = xml[:matches[0].start()]+replacement+xml[matches[0].end():]
                height_changes[str(row)] = {'Before':old_height,'After':new_height}
        with ZipFile(accepted,'w') as output:
            for entry in baseline.infolist():
                content = native.read(entry.filename) if entry.filename == 'xl/vbaProject.bin' else (
                    xml.encode('utf-8') if entry.filename == part else baseline.read(entry.filename))
                output.writestr(entry,content)
    with ZipFile(original) as baseline, ZipFile(accepted) as final:
        assert baseline.namelist() == final.namelist()
        changed_parts = [name for name in baseline.namelist() if baseline.read(name) != final.read(name)]
        assert set(changed_parts) == {'xl/vbaProject.bin',part}
    result = {'Status':'SEGMENT_NAMING_PACKAGE_PASS','BaselineSHA256':sha(original),'RegressionBaselineSHA256':sha(D/'Baseline.xlsm'),'CompiledSHA256':sha(compiled),
              'AcceptedSHA256':sha(accepted),'NativeComponents':158,'AddedComponents':[], 'ChangedComponents':sorted(changed),
              'TestedNativeMatchesCandidate':True,'BaselineProbeOnlyTestModuleChanged':True,'UserLiveVBAUnchanged':True,
              'ChangedZIPParts':changed_parts,'HelpCells':changes,'HelpRowHeights':height_changes}
    (D/'Package.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({k:v for k,v in result.items() if k != 'HelpCells'},ensure_ascii=False))


if __name__ == '__main__':
    package()
