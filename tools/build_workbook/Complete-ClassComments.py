"""Дополняет русские комментарии из назначения API и описаний полей класса.

Не меняет исполняемый VBA. Неизвестные методы требуют явного описания.
"""
from __future__ import annotations
import argparse
import importlib.util
import json
import re
from pathlib import Path
from SectionRecipeApiText import purpose
from VbeCommentText import readable

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('audit', Path(__file__).with_name('Audit-ClassComments.py'))
audit = importlib.util.module_from_spec(spec)
spec.loader.exec_module(audit)
spec = importlib.util.spec_from_file_location('api', Path(__file__).with_name('Build-SectionRecipeApi.py'))
api = importlib.util.module_from_spec(spec)
spec.loader.exec_module(api)


def property_field_meaning(lines, index):
    body = []
    for line in lines[index:]:
        body.append(line)
        if re.search(r'\bEnd Property\b', line, re.I): break
    fields = re.findall(r'\bm[A-Z]\w*\b', '\n'.join(body))
    for field in dict.fromkeys(fields):
        for i, line in enumerate(lines):
            if not re.match(r'^Private\s+(?:WithEvents\s+)?'+re.escape(field)+r'\b', line): continue
            tail = line.split("'", 1)[1].strip() if "'" in line else ''
            if audit.RUSSIAN.search(tail): return tail
            leading = audit.preceding(lines, i)
            if any(audit.RUSSIAN.search(s) for s in leading):
                return ' '.join(s.lstrip().lstrip("'").strip() for s in leading if audit.RUSSIAN.search(s))
    return None


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--apply', action='store_true')
    args = parser.parse_args()
    from ClassCommentMeanings import MEANINGS
    meanings = MEANINGS
    members = {m['id']: m for m in api.inventory()['members']}
    plans, unresolved = [], []
    files = [p for folder in ('src', 'tests', 'docs/geometry_refactor/examples') for p in (ROOT/folder).rglob('*.cls')]
    for path in sorted(files):
        content = path.read_text(encoding='utf-8-sig')
        lines = content.splitlines()
        owner = path.stem
        state = audit.inspect(path)
        additions = []
        if not state['headerRussian']:
            description = meanings['headers'].get(owner)
            if description:
                index = lines.index('Option Explicit') + 1
                additions.append((index, ['', "' ==========================================================================", "' "+description, "' =========================================================================="]))
            else: unresolved.append({'owner': owner, 'name': 'HEADER'})
        for m in state['methods']:
            if m['commented']: continue
            key = owner+'.'+m['name']
            description = meanings['methods'].get(key)
            if not description and key in members:
                description = purpose(members[key])
            if not description and m['kind'].startswith('Property'):
                description = meanings['properties'].get(key) or property_field_meaning(lines, m['line']-1) or meanings['properties'].get(m['name'])
            if not description: description = meanings['methods'].get(m['name'])
            if description:
                # Короткие физические пояснения сохраняются рядом с методом;
                # длинные исходные описания разбиваются без изменения смысла.
                import textwrap
                comments = ["' "+s for s in textwrap.wrap(readable(description), 94, break_long_words=False, break_on_hyphens=False)]
                additions.append((m['line']-1, comments))
            else:
                unresolved.append({'owner': owner, 'name': m['name'], 'kind': m['kind'], 'line': m['line'], 'body': '\n'.join(lines[m['line']-1:m['line']+5])})
        plans.append((path, lines, additions))
    report = ROOT/'docs/regression/GeometryReview_2026-10-12/CommentDescriptionsPending.json'
    report.write_text(json.dumps(unresolved, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps({'classes': len(plans), 'additions': sum(len(p[2]) for p in plans), 'unresolved': len(unresolved)}))
    if args.apply:
        if unresolved: raise SystemExit('Сначала задайте смысл всех неизвестных методов/свойств.')
        for path, lines, additions in plans:
            for index, comments in sorted(additions, reverse=True): lines[index:index] = comments
            path.write_text('\n'.join(lines)+'\n', encoding='utf-8', newline='\r\n')

if __name__ == '__main__': main()
