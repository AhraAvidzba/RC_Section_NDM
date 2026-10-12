"""Уточняет только добавленные комментарии: описание поля важнее общего имени."""
import importlib.util
import json
import textwrap
from pathlib import Path
from ClassCommentMeanings import MEANINGS
from SectionRecipeApiText import purpose
from VbeCommentText import comments_only

ROOT=Path(__file__).resolve().parents[2]
spec=importlib.util.spec_from_file_location('complete',Path(__file__).with_name('Complete-ClassComments.py'))
complete=importlib.util.module_from_spec(spec);spec.loader.exec_module(complete)

def main():
    original=json.loads((ROOT/'docs/regression/GeometryReview_2026-10-12/ClassCommentsBefore.json').read_text(encoding='utf-8'))
    members={m['id']:m for m in complete.api.inventory()['members']}
    reviewed=0
    for cls in original['details']:
        path=ROOT/cls['path'];lines=path.read_text(encoding='utf-8-sig').splitlines();owner=path.stem
        for method in cls['methods']:
            if method['commented'] or not method['kind'].startswith('Property'):continue
            key=owner+'.'+method['name']
            index=next(i for i,line in enumerate(lines) if (match:=complete.audit.METHOD.match(line)) and match[3]==method['name'] and match[2].lower()==method['kind'].lower())
            old=MEANINGS['methods'].get(key)
            if not old and key in members:old=purpose(members[key])
            if not old:old=MEANINGS['properties'].get(key) or MEANINGS['properties'].get(method['name']) or complete.property_field_meaning(lines,index)
            new=MEANINGS['methods'].get(key)
            if not new and key in members:new=purpose(members[key])
            if not new:new=MEANINGS['properties'].get(key) or complete.property_field_meaning(lines,index) or MEANINGS['properties'].get(method['name'])
            if old==new:continue
            old_lines=["' "+s for s in textwrap.wrap(old,94,break_long_words=False,break_on_hyphens=False)]
            if lines[index-len(old_lines):index]!=old_lines:raise ValueError('Generated block changed unexpectedly: '+key)
            new_lines=["' "+s for s in textwrap.wrap(new,94,break_long_words=False,break_on_hyphens=False)]
            lines[index-len(old_lines):index]=new_lines;reviewed+=1
        path.write_text(comments_only('\n'.join(lines)),encoding='utf-8',newline='\r\n')
    print(json.dumps({'status':'GENERATED_COMMENTS_REVIEWED','descriptionsRefined':reviewed}))

if __name__=='__main__':main()
