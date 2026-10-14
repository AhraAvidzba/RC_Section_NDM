"""Уточняет только пояснения; запись и нумерация не являются перемещением стержней."""
from pathlib import Path
directory=Path(__file__).resolve().parent
root=directory.parents[3]
paths=list((root/'src/Geometry').glob('CSectionRecipe*.cls'))+[root/'tests/modTestAudit02Api.bas']
publication='В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.'
circle='В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой окружности.'
side='В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой стороны.'
group='В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.'
count=0
for path in paths:
    original=path.read_text(encoding='utf-8-sig')
    actual=original.replace('Связанные стержни выводим рядом с основным.',publication)
    actual=actual.replace('Одна подпись для всех рядов окружности.',circle)
    actual=actual.replace('Общая подпись арматуры этой группы.',group if path.suffix=='.bas' else side)
    if actual!=original:
        count+=original.count('Связанные стержни выводим рядом с основным.')
        assert all(len(line)<=200 for line in actual.splitlines() if 'options.Publication =' in line or 'options.AnnotationGroup =' in line)
        path.write_text(actual,encoding='utf-8-sig',newline='\r\n')
assert count==13,count
print('RECIPE_COMMENTS_CLARIFIED',count)
