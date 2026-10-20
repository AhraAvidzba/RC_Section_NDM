"""Compare exact native geometry before and after W replaces the side-type selector."""
import importlib.util
import json
import io
from collections import Counter
from pathlib import Path
from zipfile import ZipFile

ROOT=Path(__file__).resolve().parents[2]
D=ROOT/'docs/regression/SharedRebarInput_2026-10-18/SectionNames_2026-10-20'
spec=importlib.util.spec_from_file_location('snapshots',Path(__file__).with_name('Compare-GeometrySnapshots.py'))
snap=importlib.util.module_from_spec(spec);spec.loader.exec_module(snap)

def snapshot(name):
    path=D/name
    if path.exists():return snap.read_snapshot(path)
    with ZipFile(D/'NativeSnapshotsAndReview.zip') as archive:
        class ArchivedText:
            def open(self,**kwargs):return io.StringIO(archive.read(name).decode('utf-8-sig'),newline='')
        return snap.read_snapshot(ArchivedText())

def run(before_name='BeforeGeometry.txt',after_name='AfterGeometry.txt',expected_cases=24,report_name='GeometryEquivalence.json'):
    before=snapshot(before_name);after=snapshot(after_name)
    old_inputs=sum(row[1]=='INPUT' for row in before);new_inputs=sum(row[1]=='INPUT' for row in after)
    # The input snapshot changes because two rows were removed. Compare the actual
    # built concrete, bars and annotations independently of the input table layout.
    before=[row for row in before if row[1]!='INPUT'];after=[row for row in after if row[1]!='INPUT']
    assert len(before)==len(after),(len(before),len(after))
    fields=0;numeric=0;comments=[];cases=set();counts=Counter()
    for old,new in zip(before,after):
        assert old[:3]==new[:3],(old[:3],new[:3])
        case,kind,entity=old[:3];cases.add(case);counts[kind]+=1
        for name,a,b in zip(snap.FIELDS[kind],old[3:],new[3:]):
            fields+=1
            if name=='comment':
                if a!=b:
                    # Only the human name of this section changed in these comments.
                    def caption(value):
                        return snap.decode(value).replace('PointedRect','Клиновидное сечение').replace('PointedSection','Клиновидное сечение').replace('скругленного прямоугольника','клиновидного сечения').replace('скруглённого прямоугольника','клиновидного сечения').replace('RectSet:', 'Набор прямоугольников:').replace('RectSet-сечения', 'Rectangles-сечения')
                    aa=caption(a);bb=caption(b)
                    assert aa==bb,(case,entity,name,aa,bb)
                    comments.append((case,kind,entity))
                continue
            if a.startswith(('d:','i:')):numeric+=1
            assert a==b,(case,kind,entity,name,a,b)
    assert len(cases)==expected_cases,len(cases)
    result={'Passed':True,'Cases':len(cases),'Records':len(before),'Fields':fields,
            'BitIdenticalNumericFields':numeric,'ChangedSectionNameComments':len(comments),'RecordsByKind':dict(counts),
            'GeometryAndBarsUnchanged':True,'InputRecordsBefore':old_inputs,'InputRecordsAfter':new_inputs,'FullOnOffRun':False}
    (D/report_name).write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False))

if __name__=='__main__':
    run()
    run('BeforeRectangles.txt','AfterRectangles.txt',18,'RectanglesEquivalence.json')
