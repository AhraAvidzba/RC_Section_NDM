"""ДЛЯ ТЕСТОВ: после actual full PASS готовит только VBA, сохраняя все прочие части пользовательских книг."""
from pathlib import Path
from zipfile import ZipFile
from datetime import datetime, timezone
import hashlib
import json
import posixpath
import subprocess
import xml.etree.ElementTree as ET

P=Path(__file__).resolve().parent
R=Path('C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM')
PS5='C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe'
NS={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p':'http://schemas.openxmlformats.org/package/2006/relationships'}

def test_hash(data):
    """ДЛЯ ТЕСТОВ: сравнивает фактические байты без нормализации чисел, XML или изображений."""
    return hashlib.sha256(data).hexdigest().upper()

def test_parts(path):
    """ДЛЯ ТЕСТОВ: проверяет уникальные части и CRC исходного ZIP до любых замен."""
    with ZipFile(path) as z:
        names=z.namelist()
        assert len(names)==len(set(names)), 'Duplicate ZIP part names'
        assert z.testzip() is None, 'Corrupt ZIP payload'
        return z.infolist(),{name:z.read(name) for name in names},z.comment

def test_documents(parts):
    """ДЛЯ ТЕСТОВ: требует совпадения CodeName документов, чтобы не отвязать VBA от пользовательских листов."""
    workbook=ET.fromstring(parts['xl/workbook.xml'])
    relation={x.attrib['Id']:x.attrib['Target'] for x in ET.fromstring(parts['xl/_rels/workbook.xml.rels'])}
    prop=workbook.find('m:workbookPr',NS)
    code=prop.attrib.get('codeName') if prop is not None else None
    sheets=[]
    for sheet in workbook.findall('m:sheets/m:sheet',NS):
        target=relation[sheet.attrib['{'+NS['r']+'}id']]
        target=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl',target))
        node=ET.fromstring(parts[target]).find('m:sheetPr',NS)
        sheets.append(node.attrib.get('codeName') if node is not None else None)
    assert all(sheets) and code and len(set(sheets))==len(sheets), 'Explicit unique document CodeNames required'
    return dict(WorkbookCodeName=code,SheetCodeNames=sorted(sheets))

def test_prepare_book(label,path,baseline,donor,documents,out):
    """ДЛЯ ТЕСТОВ: меняет только vbaProject.bin в частной копии; последняя сохранённая версия данных берётся из текущего файла."""
    before_bytes=path.read_bytes();before_hash=test_hash(before_bytes)
    infos,before,comment=test_parts(path)
    assert test_hash(before['xl/vbaProject.bin'])==baseline['VBAProjectSHA256'], 'VBA changed since pre-full fence: inspect and preserve manual code before publication'
    assert not any('vbaprojectsignature' in name.lower() or name.lower().startswith('_xmlsignatures/') for name in before), 'Signed package requires separate review; do not remove a signature'
    assert test_documents(before)==documents, 'User document CodeNames differ from accepted project'
    after=dict(before);after['xl/vbaProject.bin']=donor
    changed=[name for name in before if before[name]!=after[name]]
    assert changed==['xl/vbaProject.bin'], 'Only actual VBA project replacement is permitted'
    directory=out/label;directory.mkdir()
    target=directory/path.name
    with ZipFile(target,'w') as z:
        z.comment=comment
        for info in infos:z.writestr(info,after[info.filename])
    _,readback,_=test_parts(target)
    assert readback.keys()==before.keys()
    assert readback['xl/vbaProject.bin']==donor
    preserved=[name for name in before if name!='xl/vbaProject.bin']
    assert all(readback[name]==before[name] for name in preserved)
    assert path.read_bytes()==before_bytes, 'Original book changed during private preparation'
    return dict(Label=label,Original=str(path),OriginalSHA256=before_hash,PreFullSHA256=baseline['SHA256'],
        DataChangedByUserSinceFullStart=before_hash!=baseline['SHA256'],Prepared=str(target),PreparedSHA256=test_hash(target.read_bytes()),
        SourceVBAProjectSHA256=test_hash(donor),BeforeVBAProjectSHA256=baseline['VBAProjectSHA256'],ChangedParts=changed,
        PartCount=len(before),PreservedPartCount=len(preserved),AllOtherUncompressedPartBytesExact=True,
        BeforePartHashes={name:test_hash(before[name]) for name in before},AfterPartHashes={name:test_hash(readback[name]) for name in readback},
        UserDataAndResultsPreserved=True,Recalculated=False,NativeValidationPending=True)

def test_main():
    """ДЛЯ ТЕСТОВ: живой/неуспешный полный набор завершается до создания private outputs; опубликованные книги не трогает."""
    gate=subprocess.run([PS5,'-NoProfile','-ExecutionPolicy','Bypass','-File',str(P/'TestReadCurrentRun02Pass60.ps1')],cwd=R,capture_output=True)
    if gate.returncode:
        raise RuntimeError('Actual current full PASS/all-own/task cleanup required; no books prepared. '+gate.stderr.decode('utf-8',errors='replace'))
    proof=json.loads(gate.stdout.decode('utf-8-sig'))
    assert proof['OfficialPASS'] and proof['CoreSuites']==182 and proof['FocusedSingles']==122 and proof['FocusedRecords']==26
    source=Path(proof['Source']);assert test_hash(source.read_bytes())==proof['SourceSHA256']
    _,accepted,_=test_parts(source)
    documents=test_documents(accepted)
    fences=json.loads((P/'PreFullBooksFence39.json').read_text(encoding='utf-8-sig'))
    assert len(fences)==2
    out=P/'PreparedUserBooks62';out.mkdir(exist_ok=False)
    books=[]
    for label,baseline in zip(('Main','Heavy'),fences):
        path=Path(baseline['Path'])
        expected=R/'workbook/output/RC_Section_NDM.xlsm' if label=='Main' else R/'workbook/output/NDM_HeavyCircle_1232_30_2026-11-01/RC_Section_NDM_HeavyCircle_1232_30.xlsm'
        assert path.resolve()==expected.resolve(), 'Unexpected publication path'
        books.append(test_prepare_book(label,path,baseline,accepted['xl/vbaProject.bin'],documents,out))
    receipt=dict(Status='PRIVATE_CODE_ONLY_BOOKS_PREPARED_NATIVE_VALIDATION_REQUIRED',AtUTC=datetime.now(timezone.utc).isoformat(),
        AcceptedSourceSHA256=proof['SourceSHA256'],SourceCommit=proof['SourceCommit'],OfficialFullPass=True,
        SourceFullPassAppliesOnlyToAcceptedSHA=True,Documents=documents,Books=books,Published=False,ExcelStarted=0,Recalculated=False)
    (P/'PreparedUserBooks62.json').write_text(json.dumps(receipt,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({k:v for k,v in receipt.items() if k not in ('Books','Documents')},ensure_ascii=False))

if __name__=='__main__':test_main()
