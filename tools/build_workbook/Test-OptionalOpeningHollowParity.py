"""ДЛЯ ТЕСТОВ: точное сравнение 22 вариантов HollowRectangle с принятой книгой."""
import argparse, collections, hashlib, json
from pathlib import Path

def test_records(path):
    """Читает только физические записи; содержимое служебных ячеек не проверяется."""
    data=path.read_text(encoding='utf-8-sig')
    assert not any(line.startswith('FAIL') for line in data.splitlines()), path
    result=[]
    for line in data.splitlines():
        if not line: continue
        fields=line.split('\t')
        assert len(fields)>=4 and fields[0].startswith('HollowMatrix.'), (path,line[:120])
        if fields[1]=='INPUT':continue  # Служебные и неприменимые ячейки не являются расчётными результатами.
        result.append(line)
    return result

def test_main():
    """Требует полный состав сценариев и точное равенство всех сохранённых физических значений."""
    a=argparse.ArgumentParser();a.add_argument('baseline',type=Path);a.add_argument('candidate',type=Path);a.add_argument('report',type=Path);args=a.parse_args()
    checks=[]
    for name,scenarios in [('HollowMatrix',range(8)),('HollowEdgeMatrix',range(8,11))]:
        expected={f'HollowMatrix.{shape}.case{i}' for shape in ['HollowCentered','HollowOffset'] for i in scenarios}
        left=test_records(args.baseline/(name+'.txt'));right=test_records(args.candidate/(name+'.txt'))
        cases={line.split('\t')[0] for line in right};kind=collections.Counter(line.split('\t')[1] for line in right)
        differences=[{'Index':i,'Baseline':x,'Candidate':y} for i,(x,y) in enumerate(zip(left,right)) if x!=y]
        checks.append({'Dataset':name,'ExpectedCases':len(expected),'ActualCases':len(cases),'CaseSetExact':cases==expected,'BaselineRows':len(left),'CandidateRows':len(right),'Kinds':dict(kind),'PhysicalRecordsExact':left==right,'Differences':differences[:12], 'BaselineSHA256':hashlib.sha256((args.baseline/(name+'.txt')).read_bytes()).hexdigest(),'CandidateSHA256':hashlib.sha256((args.candidate/(name+'.txt')).read_bytes()).hexdigest()})
    passed=all(c['CaseSetExact'] and c['PhysicalRecordsExact'] and all(c['Kinds'].get(k,0)>0 for k in ['MODEL','PROPS','BOUNDS','FIBER','LAYOUT','CONCRETE','BAR','ANNOTATION','CONTOUR']) for c in checks)
    proof={'Status':'PASS' if passed else 'FAIL','Cases':sum(c['ExpectedCases'] for c in checks),'Ignored':'All INPUT records: service/inapplicable values are not constrained. All physical output records are compared exactly.','Checks':checks}
    args.report.write_text(json.dumps(proof,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps({'Status':proof['Status'],'Cases':proof['Cases'],'Rows':sum(c['CandidateRows'] for c in checks)}));assert passed
if __name__=='__main__':test_main()
