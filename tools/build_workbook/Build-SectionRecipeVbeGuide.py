"""Создаёт стандартный VBE-модуль только из комментариев текущего API."""
import importlib.util
import json
import re
import textwrap
from pathlib import Path
from SectionRecipeApiText import purpose, parameter_meaning, LIFECYCLE, ERRORS
from VbeCommentText import readable

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('api', Path(__file__).with_name('Build-SectionRecipeApi.py'))
api = importlib.util.module_from_spec(spec)
spec.loader.exec_module(api)

def build():
    inventory = api.inventory()
    lines = ['Attribute VB_Name = "modSectionRecipeApiGuide"']
    def comment(text=''):
        text=readable(text)
        lines.extend("' "+line if line else "'" for line in (textwrap.wrap(text, 104, break_long_words=False, break_on_hyphens=False) if text else ['']))
    def separator(title):
        comment('='*104); comment(title); comment('='*104)
    separator('ДЕКЛАРАТИВНОЕ ОПИСАНИЕ СЕЧЕНИЯ — ПОЛНЫЙ СПРАВОЧНИК API')
    comment('Модуль содержит только комментарии. Тела классов и методов находятся в соответствующих модулях VBE.')
    comment('Поиск: Ctrl+F по имени класса или метода. Краткий порядок подключения — на листе Справка.')
    comment('Внутренние единицы: мм, мм², Н, Н·мм, МПа; углы в радианах. Положительное N — растяжение.')
    comment('Ввод преобразует CUnitSystem; рецепт получает снимки значений, не Excel-объекты.')
    separator('ОСНОВНЫЕ ПОНЯТИЯ')
    comment('Путь (Path) — непрерывная геометрическая линия: отрезок, дуга, окружность или последовательность таких участков. Замкнутый путь может задавать границу бетона или отверстия; путь размещения арматуры задаёт линию центров стержней.')
    comment('Сегмент (Segment) — один отрезок или одна дуга в составе пути. Группа путей (PathGroup) объединяет независимые линии без соединений между ними.')
    comment('Положение вдоль линии задаётся расстоянием от её начала в мм. В именах API это расстояние может называться Station; например, RebarAtStations задаёт места стержней массивом таких расстояний.')
    comment('Правило арматуры (Rule) — описание диаметра и мест размещения стержней: по числу стержней, расстоянию между ними, отдельным местам или связи с другим рядом. Общий исполнитель по правилу создаёт физические стержни.')
    comment('ID — уникальное текстовое имя объекта внутри описания сечения. Привязка (Anchor) — именованная точка, используемая для размеров и подписей.')
    for title, mapping in (('ЖИЗНЕННЫЙ ЦИКЛ', LIFECYCLE), ('ОШИБКИ И ОГРАНИЧЕНИЯ', ERRORS)):
        separator(title)
        if isinstance(mapping, dict):
            for name, text in mapping.items(): comment(str(name)+': '+str(text))
        else:
            for text in mapping: comment(str(text))
    options = (ROOT/'src/Section/CRebarRuleOptions.cls').read_text(encoding='utf-8-sig')
    initialization = re.search(r'(?s)Private Sub Class_Initialize\(\)(.*?)End Sub', options)[1]
    defaults = dict(re.findall(r'(?m)^\s*(\w+) = ([^\r\n]+)', initialization))
    owner = None
    for member in inventory['members']:
        if member['owner'] != owner:
            owner = member['owner']; separator(owner)
        comment('-'*104); comment(member['id']); comment(purpose(member)); comment(member['signature'])
        if owner == 'CRebarRuleOptions' and member['kind'] == 'Field':
            default = defaults.get(member['name'], 'False' if member.get('resultType') == 'Boolean' else '0' if member.get('resultType') in ('Long', 'Double') else 'пустое значение VBA')
            comment('Default: '+default+'.')
        for field in member.get('fields', []): comment(field)
        for parameter in member.get('parameters', []):
            comment(parameter['signature']+' — '+parameter_meaning(member, parameter))
        comment('Код метода или объявление: '+member['owner']+' в окне проекта этой книги (Ctrl+R).')
    body = '\n'.join(lines)+'\n'
    assert all(not line.strip() or line.lstrip().startswith("'") for line in lines[1:])
    target = ROOT/'src/Common/modSectionRecipeApiGuide.bas'
    target.write_text(body, encoding='utf-8', newline='\r\n')
    print(json.dumps({'status': 'COMMENT_ONLY_VBE_API', 'members': len(inventory['members']), 'lines': len(lines), 'path': target.relative_to(ROOT).as_posix()}))

if __name__ == '__main__': build()
