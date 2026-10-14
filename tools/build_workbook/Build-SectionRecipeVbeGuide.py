"""Справочник автора геометрии, арматуры и аннотаций: только комментарии VBE."""
import importlib.util
import json
import re
import textwrap
from pathlib import Path
from SectionRecipeApiText import purpose, parameter_meaning
from SectionRecipeAuthorScope import selected_members
from VbeCommentText import readable
from RecipeArgumentText import argument_comment, signatures

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
    builder_signatures = signatures((ROOT/'src/Section/CSectionDefinitionBuilder.cls').read_text(encoding='utf-8-sig'))

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
        'Пробел и символ _ в конце строки VBA продолжают команду на следующей строке. Это знак записи кода, а не передаваемый параметр. Именованный аргумент записывается как имя:=значение.',
        'Initialize открывает одно описание, Finish завершает его. CirclePath, LinePath, RectanglePath и другие создающие команды сами сохраняют линию. Возвращённый path уже зарегистрирован; Outer, Opening и армирование получают его ID. AddPath нужен только для отдельно созданной линии.',
        'Таблицы и единицы context — независимый снимок ввода. context.Units используется для преобразования чисел; перенастраивать его нельзя. Проверка и размещение стержней выполняются общим Build до создания сетки.',
        'Команды подписей задают группу, геометрическую привязку и направление. Расстояние от линии арматуры до края бетона программа определяет по фактическому контуру; дополнительный отступ задаётся в настройках плоттера.',
    ]:
        comment(text, 2); comment()
    comment('МИНИМАЛЬНЫЙ ПРИМЕР', 2)
    comment('Полный класс CSectionRecipeMySection. Размеры заданы в мм; context требуется интерфейсу, входных таблиц в этом примере нет. Создайте Class Module с этим именем и вставьте текст:', 4)
    comment()
    minimal = (ROOT/'docs/geometry_refactor/examples/CSectionRecipeMySection.cls').read_text(encoding='utf-8-sig')
    minimal = minimal[minimal.index('Option Explicit'):]
    for line in minimal.splitlines():
        explanation = ''
        if re.search(r'\bfacade\.', line) and not line.lstrip().startswith("'"):
            explanation, _ = argument_comment(line.strip(), 'MySection', builder_signatures, {}, '')
        if explanation and len(line) + len(explanation) <= 165:
            lines.append("'       " + line + "  ' " + readable(explanation))
        else:
            if explanation:
                for note in textwrap.wrap(readable(explanation), 90, break_long_words=False, break_on_hyphens=False):
                    lines.append("'       ' " + note)
            lines.append("'       " + line if line else "'")
    comment()
    comment('В CSectionTypeRegistry.Class_Initialize добавьте RegisterRecipe "MySection", New CSectionRecipeMySection, Array(). Затем выполните UpdateSectionRecipeConfig. Пустой Array() подходит этому примеру с постоянными размерами; для размеров из Config укажите имена входных таблиц.', 4)

    comment('Ввод из Config передаётся в context общим CSectionInputSnapshotReader: table = context.Table("имя таблицы"), units = context.Units. table.PositiveLength/NonNegativeLength возвращают мм и проверяют знак; SourceReference возвращает адрес. Полный пример такого ввода — CSectionRecipeRoundedRectangle.', 4)
    comment()
    comment('ОПИСАТЕЛЬ И ПРИВЯЗКИ', 2)
    comment('CPolygonSpec объединяет исходные ID, X, Y, FilletRadius и SourceReference каждой вершины. New → AddVertex → Seal. Состояние читают Count/IsSealed, значения — VertexID/X/Y/FilletRadius/SourceReference по индексу от 1. После Seal изменение запрещено; повторный Seal безопасен. Первую вершину в конце не повторяют.', 4)
    comment('Исходная вершина до скругления может не лежать на контуре. PathAnchor задаёт точку готового сегмента, Anchor — заданную координату. Вершины не превращаются в anchors автоматически. OffsetRoundedPolygon: плюс внутрь исходного многоугольника; Offset: плюс влево от направления готовой линии. У группы независимых линий Offset отсутствует.', 4)
    comment('Aligned измеряет истинное расстояние, его линия параллельна измеряемому участку. Горизонтальный и вертикальный размеры измеряют проекции; PathLength суммирует прямые и дуги. Поворот текста остаётся прежним, 0/90 градусов.', 4)
    code = (ROOT/'tests/modTestAudit02Api.bas').read_text(encoding='utf-8-sig')
    for name, title, condition in [
        ('BuildAudit02DemoDefinition', 'ПОЛНЫЙ ПРИМЕР: СКРУГЛЕНИЯ, АРМАТУРА И НАКЛОННЫЙ РАЗМЕР',
         'Это самостоятельная функция; b/h/w и все объекты объявлены внутри. Edge.TR — реальная ненулевая наклонная грань между точками касания. Армирование учебное, без проверки достаточности.'),
        ('BuildAudit02OperationsDefinition', 'ПОЛНЫЙ ПРИМЕР: ПРИМИТИВЫ, СМЕЩЕНИЯ И ПРОЕКЦИЯ',
         'Самостоятельная функция с постоянными данными в мм: прямоугольники смежные, отверстие внутри Outer, линии независимые. RowNormal=1 направлен влево от MainAxis; луч проекции (0,1) пересекает ProjectedAxis. Примеры можно поместить в тот же класс и вызвать вместо прямоугольного описания.')]:
        comment(); comment(title, 2); comment(condition, 4); comment()
        body = re.search(r'(?ms)^Private Function '+name+r'\(.*?^End Function', code).group(0)
        for line in body.splitlines():
            lines.append("'       " + line if line else "'")
    comment(); comment('Миграция: прежние публичные массивы vertices/radii/vertexIDs заменены одним CPolygonSpec. AddCorner и циклы технической нумерации из рецептов удалены. Другого рабочего варианта этих операций нет.', 4)

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
