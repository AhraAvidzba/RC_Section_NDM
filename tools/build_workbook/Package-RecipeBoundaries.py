"""Переносит только проверенный VBA в исходную книгу, сохраняя все её данные и оформление побайтно."""
import hashlib
import importlib.util
import json
import re
from pathlib import Path
from zipfile import ZipFile

ROOT=Path(__file__).resolve().parents[2]
D=ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/RecipeBoundaries_2026-10-16'
spec=importlib.util.spec_from_file_location('presentation',Path(__file__).with_name('Verify-GuidePresentation.py'))
presentation=importlib.util.module_from_spec(spec);spec.loader.exec_module(presentation)


def normalized(body):
    return ''.join(part if part.startswith(('"',"'")) else part.lower() for part in re.findall(
        r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27\n]+|\n',body)).strip()


def package():
    before=presentation.components(D/'BaselineNativeVBA.txt');after=presentation.components(D/'CandidateNativeVBA.txt')
    assert before.keys()==after.keys() and len(before)==158
    changed={name for name in before if normalized(before[name][2])!=normalized(after[name][2])}
    assert changed=={'CSectionRecipeRoundedRectangle','CSectionRecipeHollowRectangle','modTestAudit02Api','modSectionRecipeApiGuide'},changed
    assert json.loads((D/'CandidateSourceEquality.json').read_text(encoding='utf-8-sig'))['Passed']
    assert json.loads((D/'Acceptance.json').read_text(encoding='utf-8'))['Status']=='RECIPE_BOUNDARIES_PASS'
    accepted=D/'Accepted.xlsm';assert not accepted.exists()
    with ZipFile(D/'Baseline.xlsm') as baseline,ZipFile(D/'Candidate/RecipeBoundaries.xlsm') as compiled,ZipFile(accepted,'w') as output:
        for entry in baseline.infolist():
            output.writestr(entry,compiled.read(entry.filename) if entry.filename=='xl/vbaProject.bin' else baseline.read(entry.filename))
    with ZipFile(D/'Baseline.xlsm') as baseline,ZipFile(accepted) as output:
        assert baseline.namelist()==output.namelist()
        assert [name for name in baseline.namelist() if baseline.read(name)!=output.read(name)]==['xl/vbaProject.bin']
    result={'Status':'RECIPE_BOUNDARIES_PACKAGE_PASS','NativeComponents':158,'AddedComponents':[],
            'ChangedComponents':sorted(changed),'ChangedZIPParts':['xl/vbaProject.bin'],
            'AcceptedSHA256':hashlib.sha256(accepted.read_bytes()).hexdigest(),'AllNonVBAPartsByteIdentical':True}
    (D/'Package.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(result))


if __name__=='__main__':package()
