"""Изображение актуального исходника: текст не сокращается и не переформулируется."""
import hashlib
import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

root = Path(__file__).resolve().parents[4]
directory = Path(__file__).resolve().parent
source = root/'src/Geometry/CSectionRecipeCircle.cls'
lines = source.read_text(encoding='utf-8').splitlines()
first = lines.index('Option Explicit')
font = ImageFont.truetype('C:/Windows/Fonts/consola.ttf', 18)
title_font = ImageFont.truetype('C:/Windows/Fonts/segoeui.ttf', 24)
line_height = 26
left, top = 80, 64
width = int(max(font.getlength(line) for line in lines[first:]))+left+35
height = top+line_height*(len(lines)-first)+25
picture = Image.new('RGB', (width, height), '#ffffff')
draw = ImageDraw.Draw(picture)
draw.rectangle((0, 0, width, 48), fill='#edf1f5')
draw.text((20, 9), 'CSectionRecipeCircle — актуальный исходный код', font=title_font, fill='#172b4d')
draw.rectangle((0, 48, left-12, height), fill='#f3f4f5')
for number, line in enumerate(lines[first:], first+1):
    y = top+(number-first-1)*line_height
    draw.text((10, y), str(number).rjust(3), font=font, fill='#7a8189')
    quoted = False
    comment_at = None
    i = 0
    while i < len(line):
        char = line[i]
        if char == '"':
            if quoted and i+1 < len(line) and line[i+1] == '"':
                i += 2
                continue
            quoted = not quoted
        if char == "'" and not quoted:
            comment_at = i
            break
        i += 1
    code = line if comment_at is None else line[:comment_at]
    draw.text((left, y), code, font=font, fill='#172b4d')
    if comment_at is not None:
        draw.text((left+font.getlength(code), y), line[comment_at:], font=font, fill='#08752d')
target = directory/'RecipeCircleCode.png'
picture.save(target)
receipt = {'Source': str(source), 'SourceSHA256': hashlib.sha256(source.read_bytes()).hexdigest(), 'FirstSourceLine': first+1, 'LastSourceLine': len(lines), 'NoCodeOmitted': True, 'TextFromActualSource': True, 'Image': target.name, 'ImageSize': picture.size, 'ImageSHA256': hashlib.sha256(target.read_bytes()).hexdigest()}
(directory/'RecipeCircleImage.json').write_text(json.dumps(receipt, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
print(json.dumps(receipt, ensure_ascii=False))
