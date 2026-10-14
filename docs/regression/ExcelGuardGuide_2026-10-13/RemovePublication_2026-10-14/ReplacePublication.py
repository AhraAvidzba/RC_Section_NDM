"""Эквивалентная замена тринадцати вызовов тремя существующими полями."""
import json
import re
from pathlib import Path

directory=Path(__file__).resolve().parent
root=directory.parents[3]
pattern=re.compile(r'(?m)^(\s*)(\w+)\.SetPublication ([^\r\n]+)$')
records=[]
for path in [*root.joinpath('src/Geometry').glob('CSectionRecipe*.cls'),root/'tests/modTestAudit02Api.bas']:
    original=path.read_bytes()
    newline='\r\n' if b'\r\n' in original else '\n'
    text=original.decode('utf-8-sig').replace('\r\n','\n')
    def replace(match):
        indent,owner,raw=match.groups()
        # Все текущие вызовы имеют три простых аргумента; комментарий не является аргументом.
        arguments=[part.strip() for part in raw.split("'",1)[0].split(',')]
        assert len(arguments)==3,raw
        group,publication,coincidence=arguments
        group_note='Одна подпись для всех рядов окружности.' if group=='"Circle"' else 'Общая подпись арматуры этой группы.'
        publication_note={'PublishPerSourcePosition':'Связанные стержни выводим рядом с основным.','PublishByRule':'Выводим стержни по порядку правил.'}[publication]
        coincidence_note={'KeepPublishedOwner':'При совпадении сохраняем первый опубликованный стержень.',
                          'RejectCoincidence':'Совпадение стержней считаем ошибкой.',
                          'MergeIdentical':'Объединяем одинаковые совпавшие стержни.'}[coincidence]
        rows=[f'{indent}{owner}.AnnotationGroup = {group} \' {group_note}',
              f'{indent}{owner}.Publication = {publication} \' {publication_note}',
              f'{indent}{owner}.Coincidence = {coincidence} \' {coincidence_note}']
        assert max(map(len,rows))<=200
        records.append({'Source':path.relative_to(root).as_posix(),'Arguments':arguments,'Statements':rows})
        return '\n'.join(rows)
    changed,count=pattern.subn(replace,text)
    if count: path.write_bytes(changed.replace('\n',newline).encode('utf-8'))
assert len(records)==13,len(records)
path=root/'src/Section/CRebarRuleOptions.cls'
text=path.read_text(encoding='utf-8-sig')
removed,count=re.subn(r"(?ms)^' Задаёт только группу количественной подписи[^\n]*\n' Концы[^\n]*\nPublic Sub SetPublication\(.*?^End Sub\n\n",'',text)
assert count==1
path.write_bytes(removed.replace('\n','\r\n').encode('utf-8'))
(directory/'ReplacementMap.json').write_text(json.dumps({'Passed':True,'CallsReplaced':len(records),'ValuesPreserved':True,'OneStatementAndCommentPerLine':True,'Records':records},ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print('PUBLICATION_PROPERTIES: calls=13; values unchanged; obsolete method removed')
