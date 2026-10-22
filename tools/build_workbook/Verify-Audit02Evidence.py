"""Сопоставляет реальные снимки Audit02, изображения Excel и повторные измерения сборки."""
from pathlib import Path
import hashlib
import json
import statistics
from PIL import Image, ImageChops

ROOT = Path(__file__).resolve().parents[2]
S = ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14'
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()

def compare():
    rows = []
    for before, after in [('BaselineGeometry','UnifiedGeometry'), ('BaselineCalculationsSP35','UnifiedCalculationsSP35'),
                          ('BaselineCalculationsSP63','UnifiedCalculationsSP63'), ('BaselinePreview','UnifiedPreview')]:
        for p in sorted((S/before).glob('modTest*.txt')):
            q = S/after/p.name
            rows.append({'Before':p.relative_to(S).as_posix(),'After':q.relative_to(S).as_posix(),
                         'Equal':p.read_bytes()==q.read_bytes(),'BeforeSHA256':sha(p),'AfterSHA256':sha(q)})
    images = []
    for p in sorted((S/'BaselinePreview').glob('*.png')):
        q = S/'UnifiedPreview'/p.name
        with Image.open(p) as a, Image.open(q) as b:
            same = a.size==b.size and ImageChops.difference(a.convert('RGB'),b.convert('RGB')).getbbox() is None
            images.append({'Image':p.name,'EqualPixels':same,'Size':a.size,'BeforeSHA256':sha(p),'AfterSHA256':sha(q)})
    assert len(images)==9
    result = {'Passed':all(r['Equal'] for r in rows) and all(r['EqualPixels'] for r in images),
              'Snapshots':rows,'Images':images,'AlignedPreview':'UnifiedPreview/Audit02Demo.png',
              'FullOnOffSuiteRun':False,'Comparison':'All rows and pixels; no exclusions or tolerance changes'}
    (S/'RegressionEquality.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    def times(folder):
        result = {}
        p = S/folder/'modTestAudit02Snapshots.TestDumpAudit02BuildPerformance.txt'
        for row in p.read_text(encoding='utf-8-sig').splitlines():
            fixture, repeat, seconds, concrete, rebar = row.split('|')
            result.setdefault(fixture,[]).append(float(seconds.replace(',','.')))
        return result
    a,b = times('BaselinePerformance'),times('UnifiedPerformanceIsolated')
    performance = {'Scope':'Full BuildWorkbookSectionModel: input snapshot, recipe, region, rules, mesh, contours and annotations. No solver/output; one warmup then five timed repeats per fixture.',
                   'Claim':'No guaranteed acceleration; machine timings are observational.', 'Fixtures':{}}
    for fixture, values in a.items():
        assert len(values)==len(b[fixture])==5
        performance['Fixtures'][fixture] = {'BeforeMedian':statistics.median(values),'AfterMedian':statistics.median(b[fixture]),
            'BeforeRange':[min(values),max(values)],'AfterRange':[min(b[fixture]),max(b[fixture])],
            'RatioMedian':statistics.median(b[fixture])/statistics.median(values)}
    (S/'Performance.json').write_text(json.dumps(performance,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({'Passed':result['Passed'],'Snapshots':len(rows),'PixelEqualImages':sum(r['EqualPixels'] for r in images),'Performance':performance['Fixtures']},ensure_ascii=False))
    if not result['Passed']: raise SystemExit(1)

if __name__=='__main__':compare()
