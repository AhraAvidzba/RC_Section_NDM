import json
import struct
from collections import defaultdict
from pathlib import Path

D = Path(__file__).parent

def number(field):
    return struct.unpack('>d',bytes.fromhex(field[2:]))[0]

out = {}
for name in ['DumpRoundedRecipeMatrix','DumpHollowRecipeMatrix','DumpHollowRecipeEdgeMatrix']:
    before = (D/'BaselineMatrices'/(name+'.txt')).read_text(encoding='utf-8-sig').splitlines()
    after = (D/'CandidateMatrices'/(name+'.txt')).read_text(encoding='utf-8-sig').splitlines()
    groups = defaultdict(lambda:{'rows':0,'maxAbsolute':0,'maxScaled':0,'strings':[],'examples':[]})
    for a,b in zip(before,after):
        aa,bb=a.split('\t'),b.split('\t')
        if a==b or aa[1]=='CONTOUR':continue
        key=aa[1];group=groups[key];group['rows']+=1
        for i,(x,y) in enumerate(zip(aa,bb)):
            if x==y:continue
            if x.startswith('d:') and y.startswith('d:'):
                v,w=number(x),number(y);delta=abs(v-w);scaled=delta/max(1,abs(v),abs(w))
                if scaled>group['maxScaled']:
                    group['maxScaled']=scaled;group['examples']=[aa[:3]+[i,v,w]]
                group['maxAbsolute']=max(group['maxAbsolute'],delta)
            elif len(group['strings'])<3:group['strings'].append([i,x,y])
    out[name]={'beforeRows':len(before),'afterRows':len(after),'differences':dict(groups)}
(D/'InitialDifferences.json').write_text(json.dumps(out,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(out,ensure_ascii=False))
