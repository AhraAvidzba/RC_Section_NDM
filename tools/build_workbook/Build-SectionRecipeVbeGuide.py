"""Справочник автора геометрии, арматуры и аннотаций: только комментарии VBE."""
import importlib.util
import json
import re
import textwrap
from pathlib import Path
from SectionRecipeApiText import purpose, parameter_meaning
from SectionRecipeAuthorScope import selected_members
from VbeCommentText import readable

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('api', Path(__file__).with_name('Build-SectionRecipeApi.py'))
api = importlib.util.module_from_spec(spec)
spec.loader.exec_module(api)


def author_text(text):
    # Описание для автора сечения; технические имена API сохраняются.
    replacements = {
        'Интерфейс единственного класса типа. Возвращает опубликованное definition через общий фасад, без сетки/баров/Shapes в рецепте.':
            'Обязательный метод класса нового типа. Возвращает завершённое описание сечения. Сетку, стержни и рисунок программа строит по этому описанию автоматически.',
        'Подготовленный снимок CSectionRecipeContext, не live Excel.':
            'Объект с копиями входных таблиц и системой единиц; ячейки Excel из рецепта не читаются.',
        'Имя корневого блока snapshot.': 'Имя входной таблицы.',
        'Число строк неизменяемого input snapshot.': 'Число строк в копии входной таблицы.',
        'Число столбцов input snapshot.': 'Число столбцов в копии входной таблицы.',
        'Читает scalar snapshot; ошибка Excel в ячейке даёт адресное исключение.':
            'Возвращает исходное значение из копии таблицы. При ошибке Excel сообщает адрес ячейки, которую нужно исправить.',
        'Явный fallback конкретного вызова. Пропущенный Optional Variant без default означает обязательное поле.':
            'Значение для пустой ячейки. В Text и Number без этого параметра пустая ячейка считается ошибкой; остальные методы используют указанное в объявлении значение по умолчанию.',
        'Nothing создаёт реальные default options.': 'Nothing создаёт условия размещения по умолчанию.',
        'исходный immutable snapshot не меняется': 'исходная геометрия не меняется',
        'Смысловой ID сущности для SourceMap.': 'Имя объекта, которому назначается ссылка на исходные данные.',
        'Возвращает последнюю диагностику query фасада с SourceMap. Nothing означает отсутствие записанного отказа.':
            'Возвращает последнюю геометрическую ошибку с адресом исходных данных. Если сохранённой ошибки нет, возвращает Nothing.',
        'Реальный адрес/программный источник, добавленный SourceMap или правилом.':
            'Адрес исходной ячейки либо пояснение, откуда взят параметр геометрии или арматуры.',
    }
    for original, replacement in replacements.items():
        text = text.replace(original, replacement)
    return text


def build():
    inventory = api.inventory()
    sections = selected_members(inventory)
    lines = ['Attribute VB_Name = "modSectionRecipeApiGuide"']

    def comment(text='', indent=0):
        prefix = "' " + ' ' * indent
        wrapped = textwrap.wrap(readable(text), 102 - indent, break_long_words=False, break_on_hyphens=False)
        lines.extend(prefix + line for line in wrapped) if text else lines.append("'")

    def heading(title):
        comment(); comment()
        comment('=' * 100)
        comment(title)
        comment('=' * 100)
        comment()

    heading('СОЗДАНИЕ СЕЧЕНИЯ: ГЕОМЕТРИЯ, АРМАТУРА И АННОТАЦИИ')
    comment('Здесь собраны готовые методы для автора нового типа сечения. Описание задаёт, что нужно получить; программа по нему строит бетонную сетку, размещает стержни и готовит схему.', 2)
    comment()
    comment('Модуль состоит только из комментариев. Точные объявления методов приведены вместе с параметрами; код открывается в соответствующем классе через окно проекта Ctrl+R. Поиск по имени — Ctrl+F.', 2)
    comment()
    comment('СОДЕРЖАНИЕ')
    for number, title, members in sections:
        comment(number + '  ' + title, 4)
    heading('КАК ЧИТАТЬ ОПИСАНИЕ')
    for text in [
        'Все координаты, длины и диаметры внутри рецепта задаются в мм, площади — в мм^2, углы — в радианах. Значения из таблиц переводит context.Units; повторно переводить уже полученные мм не нужно.',
        'Path — геометрическая линия из отрезков и дуг. Замкнутая линия может быть границей бетона или отверстия. Линия центров стержней получается, например, смещением контура на as; вычислять координаты каждого стержня вручную не требуется.',
        'Segment — один отрезок или одна дуга. PathGroup объединяет несколько независимых линий. Station в имени метода означает расстояние от начала линии в мм.',
        'ID — уникальное текстовое имя линии, ряда, привязки или подписи. Anchor — именованная точка для размеров и пояснений. Rule — описание одного способа размещения стержней.',
        'Optional означает необязательный параметр; его значение по умолчанию указано после знака =. ByVal передаёт значение, ByRef позволяет методу вернуть результат через параметр. Nothing вместо options создаёт обычные условия размещения.',
    ]:
        comment(text, 2); comment()
    comment('МИНИМАЛЬНЫЙ ПРИМЕР', 2)
    comment('Тело ISectionRecipe_CreateDefinition в классе с Implements ISectionRecipe. Пример задаёт прямоугольник и пять верхних стержней диаметром 20 мм:', 4)
    comment()
    for line, explanation in [
        ('Dim facade As CSectionDefinitionBuilder, path As CGeometryPath',
         'Объявляем переменные для сборщика описания сечения и геометрической линии. Объекты пока не создаются.'),
        ('Dim options As CRebarRuleOptions',
         'Объявляем переменную для условий размещения арматуры. Объект пока не создаётся.'),
        ('Set facade = New CSectionDefinitionBuilder',
         'Создаём сборщик, в котором будем описывать границы бетона, размещение арматуры и аннотации.'),
        ('facade.Initialize "MySection"',
         'Начинаем новое описание сечения с именем MySection.'),
        ('Set path = facade.RectanglePath("Boundary", 0#, 0#, 600#, 1000#)',
         'Создаём замкнутый контур Boundary: прямоугольник шириной 600 мм и высотой 1000 мм с левым нижним углом (0; 0).'),
        ('facade.Outer "OUTER", path.ID',
         'Назначаем контур Boundary наружной границей бетона с именем OUTER.'),
        ('Set path = facade.LinePath("TopAxis", 50#, 950#, 550#, 950#)',
         'Создаём линию центров верхних стержней от (50; 950) до (550; 950) мм: отступ от верхней и боковых граней равен 50 мм.'),
        ('Set options = New CRebarRuleOptions',
         'Создаём условия размещения арматуры со значениями по умолчанию.'),
        ('options.AnnotationGroup = "Top"',
         'Назначаем арматуре имя группы Top, для которой программа сформирует общую подпись.'),
        ('facade.RebarByCount "Top.Main", path.ID, 5, 20#, options',
         'Задаём ряд Top.Main: пять стержней диаметром 20 мм, равномерно размещённых на линии TopAxis с заданными условиями.'),
        ('facade.RebarRowLabel "Top", "Top.Main", 0#, 1#, 50#',
         'Объявляем подпись группы Top по концам верхнего ряда с направлением вверх. Количество и диаметр программа получит по фактически созданным стержням.'),
        ('facade.HorizontalDistance "DIM_B", "B", 0#, 600#, 1000#',
         'Объявляем горизонтальный размер B = 600 мм между X = 0 и X = 600 на уровне Y = 1000; отступ размерной линии задают настройки плоттера.'),
        ('Set ISectionRecipe_CreateDefinition = facade.Finish',
         'Проверяем и завершаем описание сечения, затем возвращаем его общему сборщику. Сетку, стержни и схему программа построит автоматически.'),
    ]:
        comment(line, 6)
        for explanation_line in textwrap.wrap(readable(explanation), 94, break_long_words=False, break_on_hyphens=False):
            comment("' " + explanation_line, 6)
        comment()
    comment()
    comment('В CSectionTypeRegistry.Class_Initialize добавьте RegisterRecipe "MySection", New CSectionRecipeMySection, Array(). Затем выполните UpdateSectionRecipeConfig. Пустой Array() подходит этому примеру с постоянными размерами; для размеров из Config укажите имена входных таблиц.', 4)

    options = (ROOT / 'src/Section/CRebarRuleOptions.cls').read_text(encoding='utf-8-sig')
    initialization = re.search(r'(?s)Private Sub Class_Initialize\(\)(.*?)End Sub', options)[1]
    defaults = dict(re.findall(r'(?m)^\s*(\w+) = ([^\r\n]+)', initialization))
    notes = {
        '04': 'Сначала задайте контур, затем объявите смещение на нужное as. Для второго и третьего независимого ряда задайте свои линии и свои as; связь с первым рядом не обязательна.',
        '05': 'Нулевой диаметр или нулевое число стержней отключают ряд. Пустой диаметр из таблицы читает OptionalDiameter как 0; одиночный дефис недопустим. Для каждого независимого ряда используется свой ID.',
        '06': 'Создайте New CRebarRuleOptions и измените только нужные свойства. Метод размещения копирует их при объявлении ряда, поэтому позднейшая правка options не меняет уже объявленные ряды.',
        '07': 'Для общей подписи стороны задайте одинаковую AnnotationGroup всем её рядам. Подпись считает фактически созданные стержни всей группы и исчезает, если их нет. Смещение подписи в настройках плоттера отсчитывается от границы бетона.',
        '08': 'Обычно достаточно готовых методов разделов 03–04. Для собственной последовательности отрезков и дуг: Initialize, AddLine/AddArc, Seal, затем facade.AddPath. После Seal менять состав линии нельзя.',
        '09': 'Source связывает имя геометрического объекта с адресом исходных данных. Ошибка должна указывать, какой размер или ряд нужно исправить. LastGeometryDiagnostic может вернуть Nothing, если сохранённой ошибки нет.',
    }
    all_members = []
    for number, title, members in sections:
        heading(number + '  ' + title.upper())
        if number in notes:
            comment(notes[number], 2); comment()
        owner = None
        for member in members:
            if member['owner'] != owner:
                owner = member['owner']
                comment(); comment('Класс / модуль: ' + owner, 2); comment()
            comment('-' * 94, 4)
            comment(member['id'], 4)
            comment()
            comment(author_text(purpose(member)), 6)
            comment()
            comment('Объявление:', 6)
            comment(member['signature'], 8)
            if owner == 'CRebarRuleOptions' and member['kind'] == 'Field':
                default = defaults.get(member['name'], 'False' if member.get('resultType') == 'Boolean' else '0' if member.get('resultType') in ('Long', 'Double') else 'пустое значение VBA')
                comment('По умолчанию: ' + default + '.', 6)
            if member.get('fields'):
                comment()
                comment('Допустимые значения / состав:', 6)
                for field in member['fields']:
                    comment(field, 8)
            if member.get('parameters'):
                comment()
                comment('Параметры:', 6)
                for parameter in member['parameters']:
                    comment(parameter['signature'], 8)
                    comment(author_text(parameter_meaning(member, parameter)), 10)
                    comment()
            comment()
            all_members.append(member)
    assert all(not line.strip() or line.lstrip().startswith("'") for line in lines[1:])
    target = ROOT / 'src/Common/modSectionRecipeApiGuide.bas'
    target.write_text('\n'.join(lines) + '\n', encoding='utf-8', newline='\r\n')
    manifest = {'scope': 'Методы автора геометрии, арматуры и аннотаций', 'sections': [{'number': n, 'title': t, 'members': [m['id'] for m in ms]} for n, t, ms in sections], 'members': all_members}
    (ROOT / 'docs/geometry_refactor/SECTION_RECIPE_AUTHOR_MANIFEST.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'status': 'COMMENT_ONLY_VBE_API', 'members': len(all_members), 'sections': len(sections), 'lines': len(lines)}))


if __name__ == '__main__':
    build()
