"""Уточняет комментарии изменённых команд Audit02; исполняемый текст проверяется до записи."""
from pathlib import Path
import re
import textwrap
from RecipeArgumentText import MEANINGS, HELPERS, argument_comment, signatures

ROOT = Path(__file__).resolve().parents[2]
MEANINGS['Outer'].update(namingPolicy='способ публикации имён участков', namingPrefix='префикс технических имён данной границы')
MEANINGS['Opening'].update(namingPolicy='способ публикации имён участков', namingPrefix='независимый префикс отверстия')
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

def main():
    sig = signatures((ROOT/'src/Section/CSectionDefinitionBuilder.cls').read_text(encoding='utf-8-sig'))
    updated = 0
    files = [* (ROOT/'src/Geometry').glob('CSectionRecipe*.cls'), ROOT/'docs/geometry_refactor/examples/CSectionRecipeRoundedTrapezoid.cls']
    pattern = r'\bfacade\.(?:RoundedPolygon|OffsetRoundedPolygon|Outer|Opening)\b|\b(?:AdditionalOptions|ExtraOptions)\('
    for path in files:
        text = path.read_text(encoding='utf-8-sig'); lines = text.splitlines(); helpers = signatures(text)
        for i in range(len(lines)-1,-1,-1):
            if lines[i].lstrip().startswith(("'", 'Private ', 'Public ')) or not re.search(pattern, lines[i]): continue
            code = strip_comment(lines[i]); first = i
            while first > 0 and lines[first-1].lstrip().startswith("'") and lines[first-1].strip() != "'": first -= 1
            notes = '\n'.join(lines[first:i])
            if ' — ' not in notes: first = i
            description,_ = argument_comment(code.strip(),path.stem,sig,helpers,'Подготовка геометрии или условий ряда.')
            indent = code[:len(code)-len(code.lstrip())]
            if len(code)+len(description) <= 200:
                replacement = [code+"  ' "+description]
            else:
                replacement = [indent+"' "+s for s in textwrap.wrap(description,100-len(indent),break_long_words=False,break_on_hyphens=False)]+[code]
            lines[first:i+1] = replacement; updated += 1
        after = '\n'.join(lines)+'\n'
        assert executable(text) == executable(after), path
        path.write_text(after,encoding='utf-8',newline='\r\n')
    print('Updated Audit02 command comments:',updated)

if __name__ == '__main__': main()
