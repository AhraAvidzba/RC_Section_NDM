"""Приёмка прямого использования контуров: геометрия, расчёты, схемы и удалённые промежуточные пути."""
import argparse
import hashlib
import json
import math
import re
import struct
from collections import defaultdict
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
D = ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/RecipeBoundaries_2026-10-16'
MATRICES = ['DumpCircleRecipeMatrix','DumpRectSetRecipeMatrix','DumpRoundedRecipeMatrix',
            'DumpHollowRecipeMatrix','DumpHollowRecipeEdgeMatrix']


def data(name):
    if (D/name).exists():return (D/name).read_bytes()
    with ZipFile(D/'Evidence.zip') as archive:return archive.read(name)


def text(name):return data(name).decode('utf-8-sig').replace('\r\n','\n')
def receipt(name):return json.loads(text(name))
def number(value):return struct.unpack('>d',bytes.fromhex(value.removeprefix('d:')))[0]


def close(a,b,scale=1.,tolerance=2e-12):
    assert abs(a-b) <= tolerance*max(scale,abs(a),abs(b)), (a,b,scale)


def matrices():
    results=[]
    for name in MATRICES:
        a=text('BaselineMatrices/'+name+'.txt').splitlines();b=text('CandidateMatrices/'+name+'.txt').splitlines()
        assert len(a)==len(b) and not any(line.startswith(('FAIL','ERROR:')) for line in b)
        before,after=defaultdict(list),defaultdict(list)
        max_coordinate=0.;max_property=0.;half_turns=0;renamed=0
        for x,y in zip(a,b):
            xx,yy=x.split('\t'),y.split('\t')
            if len(xx)>1 and xx[1]=='CONTOUR':
                for row,dest in [(xx,before),(yy,after)]:
                    dest[(row[0],*row[3:6],*row[14:])].append(row)
                continue
            if x==y:continue
            assert xx[:3]==yy[:3] and len(xx)==len(yy), (name,xx[:3],yy[:3])
            assert xx[1] in ['PROPS','BOUNDS','CONCRETE','FIBER'], (name,xx[1])
            area=max(1.,abs(number(xx[3]))) if xx[1]=='PROPS' else 1.
            for i,(v,w) in enumerate(zip(xx,yy)):
                if v==w:continue
                assert v.startswith('d:') and w.startswith('d:'), (name,i,v,w)
                u,z=number(v),number(w)
                if xx[1]=='PROPS':
                    # Размерные масштабы: A, A*L, L, A*L² и L; L=1000 мм охватывает fixtures.
                    if i==16:
                        close(math.remainder(u-z,math.pi),0.)
                        half_turns+=int(abs(u-z)>1.)
                        continue
                    scale=area if i==3 else area*1000. if i in [4,5] else area*1000000. if 8<=i<=15 else 1000.
                    close(u,z,scale);max_property=max(max_property,abs(u-z)/max(scale,abs(u),abs(z)))
                else:
                    if xx[1] in ['CONCRETE','FIBER']:assert i in [3,4], (name,xx[1],i)
                    assert abs(u-z)<=1e-10,(name,xx[:3],i,u,z)
                    max_coordinate=max(max_coordinate,abs(u-z))
        assert before.keys()==after.keys()
        for key,rows in before.items():
            pool=after[key].copy()
            for row in rows:
                found=next((j for j,candidate in enumerate(pool) if all(abs(number(v)-number(w))<=1e-10 for v,w in zip(row[6:14],candidate[6:14]))),None)
                assert found is not None, (name,key,row)
                candidate=pool.pop(found);renamed+=int(row[2]!=candidate[2])
                max_coordinate=max(max_coordinate,max(abs(number(v)-number(w)) for v,w in zip(row[6:14],candidate[6:14])))
            assert not pool
        results.append({'Matrix':name,'Rows':len(a),'Exact':a==b,'MaxCoordinateDifferenceMM':max_coordinate,
                        'MaxScaledPropertyDifference':max_property,'EquivalentAxisHalfTurns':half_turns,'RenumberedContourSegments':renamed})
    assert text('BaselineMatrices/DumpRecipeRuleContracts.txt')==text('CandidateMatrices/DumpRecipeRuleContracts.txt')
    return results


def paths(name):
    result={};case=None;path=None
    for line in text(name).splitlines():
        parts=line.split('|')
        if parts[0]=='CASE':case=tuple(parts[1:]);result[case]={}
        elif parts[0]=='PATH':path=parts[1];result[case][path]=[parts]
        else:
            assert parts[0]=='SEGMENT',line
            result[case][path].append(parts)
    return result


def path_contracts():
    a=paths('BaselineMatrices/DumpRecipePathContracts.txt');b=paths('CandidateMatrices/DumpRecipePathContracts.txt')
    assert a.keys()==b.keys() and len(a)==60
    for case,old in a.items():
        rename={}
        if 'Shape' in old:rename={'Shape':'OUTER.Boundary' if 'OUTER.Boundary' in old else 'Boundary'}
        if 'Opening.Shape' in old:rename['Opening.Shape']='OPENING.Boundary'
        assert b[case].keys()==old.keys()-rename.keys(),case
        inverse={v:k for k,v in rename.items()}
        for name,rows in b[case].items():
            original=old[inverse.get(name,name)]
            expected=[]
            for fields in original:
                mapped=fields.copy()
                if mapped[0]=='PATH':mapped[1]=name
                else:mapped[3]=rename.get(mapped[3],mapped[3])
                expected.append(mapped)
            assert rows==expected,(case,name)
    return {'Cases':len(a),'PathsBefore':sum(map(len,a.values())),'PathsAfter':sum(map(len,b.values())),
            'ExactSourceGeometryAndParameters':True}


def calculations():
    results=[]
    for code in ['SP35','SP63']:
        a=text('Baseline'+code+'/DumpAudit02Calculations.txt').splitlines()
        b=text('Candidate'+code+'/DumpAudit02Calculations.txt').splitlines()
        assert len(a)==len(b)==32
        delta=0.;statuses=0
        for x,y in zip(a,b):
            xx,yy=x.split('|'),y.split('|');assert xx[:3]==yy[:3]
            if xx[2]=='VALUES':
                for v,w in zip(xx[3:],yy[3:]):
                    u,z=number(v),number(w);close(u,z)
                    delta=max(delta,abs(u-z)/max(1.,abs(u),abs(z)))
            else:
                assert x==y,(code,x,y)
                statuses+=int(xx[2]=='STATUS')
        assert statuses==8
        results.append({'Code':code,'Combinations':statuses,'StatusesAndCommentsExact':True,'MaxScaledValueDifference':delta})
    return results


def plots():
    from PIL import Image,ImageChops
    from io import BytesIO
    results=[]
    for name in ['Circle','Rectangle','LSection','TwoRectangles','RoundedSimple','RoundedTapered',
                 'RoundedMixed','RoundedMixedLeft','HollowCentered','HollowOffset']:
        a=data('BaselinePlots/'+name+'.png');b=data('CandidatePlots/'+name+'.png')
        pixels=[]
        if a!=b:
            diff=ImageChops.difference(Image.open(BytesIO(a)).convert('RGB'),Image.open(BytesIO(b)).convert('RGB'))
            pixels=[value for value in diff.get_flattened_data() if max(value)>0]
            assert name=='RoundedMixedLeft' and len(pixels)<=3 and max(map(max,pixels))<=1
        results.append({'Name':name,'ExactPNG':a==b,'ChangedPixels':len(pixels),'MaxChannelDifference':max(map(max,pixels),default=0)})
    return results


def scope():
    allowed={'src/Geometry/CSectionRecipePointedRect.cls','src/Geometry/CSectionRecipeHollowRectangle.cls','src/Common/modSectionRecipeApiGuide.bas'}
    files={path.relative_to(ROOT).as_posix():path for path in (ROOT/'src').rglob('*') if path.suffix in ['.bas','.cls']}
    with ZipFile(D/'BaselineSource.zip') as archive:
        old={name for name in archive.namelist() if name.startswith('src/') and name.endswith(('.bas','.cls'))}
        assert old==files.keys()
        changes={name for name,path in files.items() if path.read_text(encoding='utf-8-sig')!=archive.read(name).decode('utf-8-sig').replace('\r\n','\n')}
    assert changes==allowed,changes
    for path in (ROOT/'src/Geometry').glob('CSectionRecipe*.cls'):
        body=path.read_text(encoding='utf-8-sig').replace(' _\n',' ')
        assert not re.search(r'facade\.Chain\([^\n]*True,\s*True\)',body),path
    return {'ChangedRuntimeClasses':sorted(allowed-{'src/Common/modSectionRecipeApiGuide.bas'}),
            'KernelSolverAndPlacementCodeUnchanged':True,'SourceComponentsUnchanged':True}


def verify(published):
    for stage in ['BaselineMatrices','CandidateMatrices','Directed','BaselinePlots','CandidatePlots','BaselineSP35','BaselineSP63','CandidateSP35','CandidateSP63']:
        run=receipt(stage+'/Run.json');assert run['Passed'] and run['ProductionUnchanged'] and all(item['Passed'] for item in run['Runs']),stage
    passed=0
    for run in receipt('Directed/Run.json')['Runs']:
        totals=re.findall(r'(?m)^TOTAL[^\n]*passed=(\d+); failed=(\d+)',text('Directed/'+run['Macro']+'.txt'))
        assert len(totals)==1 and totals[0][1]=='0'
        passed+=int(totals[0][0])
    result={'Status':'RECIPE_BOUNDARIES_PASS','Suites':8,'Passed':passed,'Failed':0,'SourceScope':scope(),
            'Matrices':matrices(),'Paths':path_contracts(),'Calculations':calculations(),'Plots':plots(),'FullOnOffRun':False,'PublishedVerified':published}
    before=receipt('Before.json')
    report=ROOT/'workbook/output/RC_Section_NDM_execution_report.txt'
    assert hashlib.sha256(report.read_bytes()).hexdigest().upper()==before['Hashes']['workbook/output/RC_Section_NDM_execution_report.txt']
    if published:
        assert receipt('Publication.json')['Status']=='PUBLISHED_VERIFIED'
        for name in ['PublishedSourceEquality.json','PublishedGuidePresentation.json']:assert receipt(name)['Passed']
        assert (ROOT/'workbook/output/VBA_All_Code.txt').read_bytes()==data('PublishedNativeVBA.txt')
        assert hashlib.sha256((ROOT/'workbook/output/RC_Section_NDM.xlsm').read_bytes()).hexdigest()==receipt('Publication.json')['WorkbookSHA256']
        sources={path.stem:path for folder in ['src','tests'] for path in (ROOT/folder).rglob('*') if path.suffix in ['.cls','.bas']}
        equality=receipt('PublishedSourceEquality.json');assert equality['SourceComponents']==153 and equality['ExportComponents']==158
        for row in equality['Components']:assert hashlib.sha256(sources[row['Name']].read_bytes()).hexdigest().upper()==row['SourceSHA256']
    (D/'Acceptance.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({key:result[key] for key in ['Status','Suites','Passed','Failed','Paths','Calculations','PublishedVerified']},ensure_ascii=False))


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--published',action='store_true')
    verify(parser.parse_args().published)
