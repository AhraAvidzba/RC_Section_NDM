"""Фиксирует проверку ожиданий, состава и входов перед полным ON/OFF-прогоном."""
from pathlib import Path
from zipfile import ZipFile
import argparse,hashlib,json,re
import xml.etree.ElementTree as E

ROOT=Path(__file__).resolve().parents[2]


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory',type=Path,required=True)
    args=parser.parse_args();base=args.directory
    read=lambda name:json.loads((base/name).read_text('utf-8-sig'))
    preflight=read('Checks01/Preflight.json');package=read('Package.json')
    assert preflight['Status']=='PASSED' and preflight['SourceUnchanged']
    assert all(r['Passed'] and r['ExitCode']==0 for r in preflight['Results'])
    assert preflight['SourceSHA256']==package['AcceptedSHA256']==sha(base/'Accepted.xlsm')
    logs=[]
    for result in preflight['Results']:
        p=ROOT/result['Report'];text=p.read_text('utf-8-sig')
        started=re.findall(r'(?m)^SUITE_STARTED: ([^;\r\n]+)',text)
        finished=re.findall(r'(?m)^SUITE_FINISHED: ([^;\r\n]+)',text)
        assert started==finished==result['Macros']
        assert 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' in text
        assert not re.search(r'(?m)^(FAIL:|RUNTIME ERROR|SCRIPT ERROR|CLEANUP ERROR)',text)
        assert not re.search(r'failed=[1-9][0-9]*',text)
        logs.append({'Report':result['Report'],'SHA256':sha(p),'Suites':len(started)})
    focus=read('FocusedManifest/InputManifest.json')
    assert focus['Passed'] and focus['SingleCasesTwoModes']==122
    assert focus['SourceSHA256']==package['AcceptedSHA256']
    assert sorted(g['CasesPerMode'] for g in focus['Groups'])==[6,8,8,8,8,23]
    assert read('SourceEquality.json')['Passed']
    material=(ROOT/'tests/modTestMaterialDiagrams.bas').read_text('utf-8-sig')
    pattern=re.search(r'pattern.Pattern = "(.*)"',material).group(1).replace('""','"')
    grammar=re.compile(pattern)
    for formula in ('=$I$33','=-Config!$I$33',"=-'Config'!$I$33",'=IF($B$8=0,0,Config!$B$8*$I$30)'):
        assert grammar.fullmatch(formula),formula
    for formula in ('=SUM(A1:A3)','=INDIRECT("I33")','=Other!$I$33',"='[Other.xlsx]Config'!$I$33"):
        assert not grammar.fullmatch(formula),formula
    n={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with ZipFile(base/'Accepted.xlsm') as z:
        workbook=E.fromstring(z.read('xl/workbook.xml'))
        rel={r.attrib['Id']:r.attrib['Target'] for r in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        sheet=next(s for s in workbook.find('m:sheets',n) if s.attrib['name']=='Config')
        target=rel[sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']]
        tree=E.fromstring(z.read(target.lstrip('/') if target.startswith('/') else 'xl/'+target))
        formulas=[]
        for c in tree.findall('.//m:c',n):
            match=re.fullmatch(r'(AB|AC)(\d+)',c.attrib['r'])
            f=c.find('m:f',n)
            if match and 3<=int(match[2])<=54 and f is not None:
                formula='='+f.text;assert grammar.fullmatch(formula),(c.attrib['r'],formula)
                formulas.append(c.attrib['r'])
        assert 0<len(formulas)<=104,len(formulas)
    result={'Passed':True,'AcceptedSHA256':package['AcceptedSHA256'],'PreflightReports':logs,
            'PreflightSuites':sum(x['Suites'] for x in logs),'FocusedCases':122,
            'MaterialControlFormulasChecked':len(formulas),'FormulaGrammarPositiveControls':4,'FormulaGrammarNegativeControls':4,
            'NumericalTolerancesChanged':False,'ProductionSourceChanged':False,
            'ReviewedChanges':['Config! in direct material control references','Human-readable extra-row caption with exact row number',
                               'Circle starts at 12 oclock clockwise','2.5 mm periodic-profile allowance',
                               'Per-section named input tables and PointedSection','Complete offset boundary and skipped-position warnings'],
            'Scope':'Preflight verifies fixtures, changed expectations, material values and geometry; complete acceptance still requires the full run.'}
    (base/'ExpectationsAudit.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(result,ensure_ascii=True))


if __name__=='__main__':
    main()
