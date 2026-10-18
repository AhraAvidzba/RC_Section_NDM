"""Сверяет физические результаты до и после разделения геометрии и армирования."""
from pathlib import Path
import argparse,json,re

p=argparse.ArgumentParser();p.add_argument('--before',required=True);p.add_argument('--after',required=True);p.add_argument('--report',required=True)
a=p.parse_args();before=Path(a.before);after=Path(a.after);result=[]
for name in ('DumpCircleRecipeMatrix','DumpRectSetRecipeMatrix','DumpRoundedRecipeMatrix','DumpHollowRecipeMatrix','DumpHollowRecipeEdgeMatrix'):
    def records(directory):
        rows=(directory/(name+'.txt')).read_text(encoding='utf-8-sig').splitlines()
        assert not any(row.startswith('FAIL:') for row in rows),name
        return [row for row in rows if '\tINPUT\t' not in row]
    old,new=records(before),records(after)
    assert old==new, f'{name}: first mismatch '+str(next(((i,x,y) for i,(x,y) in enumerate(zip(old,new)) if x!=y),('length',len(old),len(new))))
    result.append({'Name':name,'Records':len(new),'ExactDoubleBits':True})
name='DumpRecipePathContracts.txt'
def contracts(directory):
    text=(directory/name).read_text(encoding='utf-8-sig')
    assert 'FAIL:' not in text
    return re.sub(r'Config!\$?[A-Z]+\$?\d+','Config!<moved-cell>',text).splitlines()
old,new=contracts(before),contracts(after)
assert old==new, 'Path/rule mismatch: '+str(next(((i,x,y) for i,(x,y) in enumerate(zip(old,new)) if x!=y),('length',len(old),len(new))))
report={'Passed':True,'Geometry':result,'RuleAndPathRecords':len(new),'OnlySourceAddressesNormalized':True}
Path(a.report).write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(report,ensure_ascii=False))
