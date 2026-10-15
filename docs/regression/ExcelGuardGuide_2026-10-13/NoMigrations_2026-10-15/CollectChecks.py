"""Сводит только принятые адресные проверки и точные снимки конечного кандидата."""
import hashlib
import json
import re
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

directory = Path(__file__).resolve().parent
def read(name):
    return json.loads((directory/name).read_text(encoding='utf-8-sig'))
def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

snapshots = []
with ZipFile(directory/'ExactSnapshots.zip', 'w', ZIP_DEFLATED) as archive:
    for kind, macro, size_key, count in [
        ('Models', 'modTestShapeArchitecture.DumpShapeArchitectureSnapshot', 'Models', 9),
        ('NDS', 'modTestShapeArchitecture.DumpPublicationControlStates', 'ControlStates', 12),
    ]:
        stage = 'Acceptance'+kind
        run = read(stage+'/Run.json')
        assert run['Passed'] and run['ProductionUnchanged']
        assert run['SourceSHA256'].lower() == sha(directory/'Candidate.xlsm')
        paths = [directory/p/(macro+'.txt') for p in ('Baseline'+kind, stage)]
        assert paths[0].read_bytes() == paths[1].read_bytes(), kind
        for path in paths:
            archive.write(path, path.relative_to(directory).as_posix())
        snapshots.append({'Kind':kind, 'AcceptedStage':stage, 'ExactByteEquality':True,
                          'Bytes':paths[0].stat().st_size, 'SHA256':sha(paths[0]), size_key:count})
(directory/'Equivalence.json').write_text(json.dumps({'Passed':True,'Snapshots':snapshots},indent=2)+'\n',encoding='utf-8')

groups = []
for stage in ('FinalSettings','Compatibility','FinalPlotAccepted2','AddedRecipe'):
    run = read(stage+'/Run.json')
    assert run['Passed'] and run['ProductionUnchanged']
    for item in run['Runs']:
        text = (directory/stage/(item['Macro']+'.txt')).read_text(encoding='utf-8-sig')
        total = re.findall(r'^TOTAL[^\r\n]*passed=(\d+); failed=(\d+)',text,re.M)
        assert total and len(total) == 1 and int(total[0][1]) == 0, (stage,item['Macro'])
        groups.append({'Stage':stage, 'Macro':item['Macro'], 'Passed':int(total[0][0]),
                       'Failed':0, 'ReportSHA256':sha(directory/stage/(item['Macro']+'.txt'))})
result = {'Passed':True,'Groups':groups,'Total':sum(item['Passed'] for item in groups)}
assert result['Total'] == 9987, result['Total']
(directory/'DirectedCounts.json').write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'Passed':True,'FocusedAssertions':result['Total'],'Snapshots':snapshots}))
