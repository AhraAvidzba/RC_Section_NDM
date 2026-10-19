"""Сравнивает исходные биты CAD и конечные результаты; рабочие данные не округляет."""
from pathlib import Path
import argparse,json,math,struct,sys

def number(text):
    value=struct.unpack('<d',bytes.fromhex(text))[0]
    if not math.isfinite(value): raise ValueError('Non-finite value')
    return value

def geometry(a,b):
    aa,bb=[p.read_text('utf-8-sig').splitlines() for p in (a,b)]
    failures=[]; maxima={}; changed=0
    if len(aa)!=len(bb):failures.append('Different geometry row counts')
    for index,(left,right) in enumerate(zip(aa,bb),1):
        if left==right:continue
        x,y=left.split('|'),right.split('|')
        if len(x)!=12 or len(y)!=12 or x[:3]!=y[:3] or x[0]!='C':
            failures.append(f'Geometry metadata/shape changed at row {index}');continue
        vx,vy=[list(map(number,z[3:])) for z in (x,y)]
        area=max(abs(vx[2]),abs(vy[2]))
        size=max(abs(vx[6]),abs(vx[7]),abs(vy[6]),abs(vy[7]),area*(vx[3]**2+vx[4]**2)/12,1)
        # Смена представления кривой меняет порядок вычислений ядра AutoCAD.
        # Допуск инерций ограничен и оценкой ошибок вычитания A*x², и 10^-10
        # характерной местной инерции. Это проверка доказательств, не допуск решателя.
        origin=max(sum(v*v for v in z[:2])+z[3]**2+z[4]**2 for z in (vx,vy))
        inertia_limit=min(128*sys.float_info.epsilon*(area*origin+size),1e-10*size)
        limits=[1e-9,1e-9,1e-9,1e-9,1e-9,1e-10,inertia_limit,inertia_limit,inertia_limit]
        names=['X_mm','Y_mm','Area_mm2','Width_mm','Height_mm','Angle_rad','Ix_mm4','Iy_mm4','Ixy_mm4']
        for field,u,v,limit in zip(names,vx,vy,limits):
            difference=abs(u-v)
            if difference:
                changed+=1
                maxima[field]=max(maxima.get(field,0),difference)
                if difference>limit:failures.append(f'{x[1]} {field}: {difference} > {limit}')
    return {'Rows':len(aa),'DifferentNumbers':changed,'MaxAbsoluteDifferences':maxima,'Failures':failures}

def calculation(a,b,section_size):
    def read(p):
        result={}
        for row in p.read_text('utf-8-sig').splitlines():
            if not row.strip():continue
            x=row.split('|',3)
            if len(x)!=4:raise ValueError(f'Invalid calculation row {row}')
            key=tuple(x[:2])
            if key in result:raise ValueError(f'Duplicate field {key}')
            result[key]=x[2:]
        return result
    aa,bb=read(a),read(b);failures=[];maximum={};changed=0;statuses=0
    if aa.keys()!=bb.keys():failures.append('Different calculation fields/states')
    if aa.get(('Control','Cases'))!=['T','6'] or len(aa)<300:failures.append('Incomplete control matrix')
    scales={}
    for source in (aa,bb):
        for key,value in source.items():
            if value[0]!='D':continue
            case=key[0].split('.')[0];s=scales.setdefault(case,{'N':1.,'M':1.})
            if key[1] in ('TargetN','Nint','NUltimate'):s['N']=max(s['N'],abs(number(value[1])))
            if key[1] in ('TargetMx','TargetMy','Mxint','Myint','MxUltimate','MyUltimate'):s['M']=max(s['M'],abs(number(value[1])))
    for key,x in aa.items():
        y=bb.get(key)
        if y is None:continue
        if x[0]!=y[0]:failures.append(f'Type changed {key}');continue
        if x[0]=='T':
            statuses+=1
            if x!=y:failures.append(f'Status/applicability changed {key}: {x} -> {y}')
        else:
            u,v=number(x[1]),number(y[1]);difference=abs(u-v)
            # Нулевой компонент равновесия сравниваем с масштабом всего вектора
            # нагрузки: относительная ошибка около нуля не имеет физического смысла.
            s=scales[key[0].split('.')[0]]
            scale=max(abs(u),abs(v))
            if key[1] in ('Nint','NUltimate','TargetN'):scale=max(scale,s['N'],s['M']/section_size)
            if key[1] in ('Mxint','Myint','MxUltimate','MyUltimate','TargetMx','TargetMy'):scale=max(scale,s['M'],s['N']*section_size)
            limit=1e-12+1e-7*scale
            if difference:
                changed+=1
                field=key[1]
                maximum[field]=max(maximum.get(field,0),difference)
                if difference>limit:failures.append(f'Calculation {key}: {difference} > {limit}; {u} -> {v}')
    return {'Fields':len(aa),'ExactStatusFields':statuses,'DifferentNumbers':changed,'MaxAbsoluteDifferences':maximum,'Failures':failures}

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--directory',type=Path,required=True)
    args=p.parse_args();d=args.directory
    points=[row.split('|') for row in (d/'A.geometry.txt').read_text('utf-8-sig').splitlines() if row.startswith('C|')]
    section_size=max(max(number(p[k]) for p in points)-min(number(p[k]) for p in points) for k in (3,4))
    if section_size<=0:raise ValueError('Degenerate control section')
    result={'Geometry':geometry(d/'A.geometry.txt',d/'B.geometry.txt'),'Calculations':{},'CalculationRelativeTolerance':1e-7,'CalculationAbsoluteFloor':1e-12,'ForceMomentScaling':'Whole load vector and section size, including components close to zero.'}
    for mode in ('Yes','No'):
        result['Calculations'][mode]=calculation(d/f'A.{mode}.calc.txt',d/f'B.{mode}.calc.txt',section_size)
    result['Passed']=not(result['Geometry']['Failures'] or any(v['Failures'] for v in result['Calculations'].values()))
    result['ToleranceScope']='Evidence comparison only; no changes to model data, cache keys, solver or normative coefficients.'
    (d/'Comparison.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(result,ensure_ascii=True))
    raise SystemExit(0 if result['Passed'] else 1)
