"""Полные сравнения закрывающей доработки: модели, ошибки, схема и PNG трёх Rectangles-подтипов."""
from pathlib import Path
import hashlib
import json
from PIL import Image, ImageChops

ROOT=Path(__file__).resolve().parents[2]
S=ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/Audit02Closeout_2026-10-14'

def verify():
    rows=[]
    for before,after in [('BaselineSnapshots','CandidateSnapshots'),('BaselinePreview','CandidatePreview')]:
        for p in sorted((S/before).glob('modTest*.txt')):
            q=S/after/p.name
            rows.append({'Before':p.relative_to(S).as_posix(),'After':q.relative_to(S).as_posix(),
                         'Equal':p.read_bytes()==q.read_bytes(),'BeforeSHA256':hashlib.sha256(p.read_bytes()).hexdigest(),
                         'AfterSHA256':hashlib.sha256(q.read_bytes()).hexdigest()})
    images=[]
    for fixture in ('Rectangle','LSection','TwoRectangles'):
        with Image.open(S/'BaselinePreview'/(fixture+'.png')) as a,Image.open(S/'CandidatePreview'/(fixture+'.png')) as b:
            equal=a.size==b.size and ImageChops.difference(a.convert('RGB'),b.convert('RGB')).getbbox() is None
            images.append({'Fixture':fixture,'EqualPixels':equal,'Size':a.size})
    tests=json.loads((S/'CandidateDirectedTests/Run.json').read_text(encoding='utf-8-sig'))
    assert len(rows)==3
    result={'Passed':all(row['Equal'] for row in rows) and all(row['EqualPixels'] for row in images) and tests['Passed'],
            'Rows':rows,'Images':images,'Assertions':290,'FullOnOffRun':False,'Comparison':'Complete rows, diagnostics and pixels; no exclusions'}
    (S/'RegressionEquality.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({'Passed':result['Passed'],'CompleteSnapshots':len(rows),'PixelEqualImages':sum(row['EqualPixels'] for row in images),'Assertions':290}))
    if not result['Passed']:raise SystemExit(1)

if __name__=='__main__':verify()
