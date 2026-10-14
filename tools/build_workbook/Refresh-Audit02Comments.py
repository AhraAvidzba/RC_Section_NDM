"""Уточняет комментарии изменённых команд Audit02; исполняемый текст проверяется до записи."""
from pathlib import Path
import re
import textwrap
from RecipeArgumentText import MEANINGS, HELPERS, argument_comment, signatures

ROOT = Path(__file__).resolve().parents[2]
for helper in ('AdditionalOptions','ExtraOptions'):
    HELPERS[helper].update(baseDiameter='диаметр основного стержня, мм', diameter='диаметр добавляемого стержня, мм',
        interveningDiameterSum='сумма промежуточных номинальных диаметров, мм')
for helper in ('AxisPath','AddBoundary','AddGeometry','AddReinforcement'):
    HELPERS[helper].update(polygon='завершённые исходные вершины с радиусами',
        outerPolygon='описание наружной границы', openingPolygon='описание отверстия')

def strip_comment(line):
    return re.split(r"'(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)",line,1)[0].rstrip()

def executable(text):
    return [strip_comment(line) for line in text.splitlines() if strip_comment(line).strip()]

def statements(lines):
    """Возвращает полную команду вместе с физическими границами, сохраняя «_» в исходнике."""
    start, pending = 0, []
    for i,line in enumerate(lines):
        if not pending: start=i
        pending.append(strip_comment(line))
        if pending[-1].endswith(' _'): continue
        logical=re.sub(r'\s+_\n\s*',' ','\n'.join(pending))
        yield start,i,pending,logical
        pending=[]
    if pending: raise ValueError('Unfinished VBA continuation')

def main():
    sig = signatures((ROOT/'src/Section/CSectionDefinitionBuilder.cls').read_text(encoding='utf-8-sig'))
    updated = 0
    files = [* (ROOT/'src/Geometry').glob('CSectionRecipe*.cls'), ROOT/'docs/geometry_refactor/examples/CSectionRecipeRoundedTrapezoid.cls']
    pattern = r'\bfacade\.(?:RoundedPolygon|OffsetRoundedPolygon|Outer|Opening)\b|\b(?:AdditionalOptions|ExtraOptions|AxisPath|AddGeometry)\(|\b(?:AddGeometry|AddBoundary|AddReinforcement|AddDimensions)\s+facade\b'
    for path in files:
        text = path.read_text(encoding='utf-8-sig'); lines = text.splitlines(); helpers = signatures(text)
        for i,last,physical,code in reversed(list(statements(lines))):
            if code.lstrip().startswith(('Private ', 'Public ')) or not re.search(pattern, code): continue
            first = i
            while first > 0 and lines[first-1].lstrip().startswith("'") and lines[first-1].strip() != "'": first -= 1
            notes = '\n'.join(lines[first:i])
            if ' — ' not in notes: first = i
            description,_ = argument_comment(code.strip(),path.stem,sig,helpers,'Подготовка геометрии или условий ряда.')
            indent = code[:len(code)-len(code.lstrip())]
            if len(physical)==1 and len(code)+len(description) <= 200:
                replacement = [physical[0]+"  ' "+description]
            else:
                replacement = [indent+"' "+s for s in textwrap.wrap(description,100-len(indent),break_long_words=False,break_on_hyphens=False)]+physical
            lines[first:last+1] = replacement; updated += 1
        after = '\n'.join(lines)+'\n'
        assert executable(text) == executable(after), path
        path.write_text(after,encoding='utf-8',newline='\r\n')
    print('Updated Audit02 command comments:',updated)

if __name__ == '__main__': main()
