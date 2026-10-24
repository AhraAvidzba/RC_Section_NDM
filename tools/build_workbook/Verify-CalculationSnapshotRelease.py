"""Независимая приёмка сохранённого ввода, прежних результатов и фактического VBA книги."""
from __future__ import annotations

import argparse
import hashlib
import importlib
import json
import re
from pathlib import Path

read = importlib.import_module('Verify-SectionRecipeGuide').read


def source_body(path):
    """Отделяет экспортные атрибуты от кода; комментарий-пример Option Explicit не является началом модуля."""
    text = path.read_text('utf-8-sig')
    start = re.search(r'^Option Explicit\b', text, re.M)
    return text[start.start():] if start else re.sub(r'^Attribute[^\n]*\n', '', text, flags=re.M)


def normalized_vba(text):
    """Учитывает нормализацию регистра идентификаторов VBA, сохраняя литералы и русские комментарии."""
    result = []
    for line in text.splitlines():
        line = line.rstrip()
        if not line.strip() or re.fullmatch(r'=+', line):
            continue
        parts = re.split(r'("(?:[^"]|"")*"|\'.*)', line)
        def code(part):
            # VBE переписывает малые десятичные литералы в экспоненциальную форму.
            # VBE также ограничивает текст числовой константы 15 значащими цифрами
            # и убирает избыточный # у дробного Double. Целые, Single и Currency не меняем.
            part = re.sub(r'(?<![\w&])(?:\d+\.\d*(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?|\d+[eE][+-]?\d+)#?(?![\w@!#])',
                          lambda match: '<double:' + format(float(match[0].rstrip('#')), '.15g') + '>', part)
            return re.sub(r'\s+', '', part).casefold()
        result.append(''.join(part if part.startswith(('"', "'")) else code(part) for part in parts))
    return result


def exported_components(path):
    """Читает собственный экспорт Excel, а не принимает текст исходников за код книги."""
    text = path.read_text('utf-8-sig')
    matches = list(re.finditer(r'^COMPONENT: (.+)\nTYPE: (\d+)\nLINES: (\d+)\n=+\n', text, re.M))
    return {match[1]: (int(match[2]), text[match.end():matches[i+1].start() if i+1 < len(matches) else len(text)]) for i, match in enumerate(matches)}


def rectangle(book, name):
    """Разбирает фактический именованный прямоугольник, без жёстких адресов входных таблиц."""
    target = book['names'][name]
    match = re.fullmatch(r"'?([^!']+)'?!\$([A-Z]+)\$(\d+)(?::\$([A-Z]+)\$(\d+))?", target)
    if not match:
        raise ValueError(f'Не прямоугольный диапазон {name}: {target}')
    sheet, left, top, right, bottom = match.groups()
    def column(text):
        value = 0
        for letter in text:
            value = value * 26 + ord(letter) - 64
        return value
    return sheet, column(left), int(top), column(right or left), int(bottom or top)


def column_name(number):
    """Возвращает буквенный адрес только для независимого чтения XML."""
    text = ''
    while number:
        number, remainder = divmod(number - 1, 26)
        text = chr(65 + remainder) + text
    return text


def cell_content(book, sheet, row, column):
    """Сравнивает формулы и значения; допустимое обновление кэша формулы учитывается отдельно."""
    cell = book['sheets'][sheet]['cells'].get(f'{column_name(column)}{row}', {})
    return (cell.get('formula'), None if cell.get('formula') else cell.get('value'))


def table(book, name):
    """Возвращает строки указанного входного диапазона в относительных координатах."""
    sheet, left, top, right, bottom = rectangle(book, name)
    return [[cell_content(book, sheet, row, column) for column in range(left, right+1)] for row in range(top, bottom+1)]


def anchored_table(book, name, width):
    """Считывает техническую таблицу до первой незаполненной строки её собственного столбца."""
    sheet, left, top, *_ = rectangle(book, name)
    rows = []
    for row in range(top, 1048577):
        if cell_content(book, sheet, row, left) == (None, None):
            break
        rows.append([cell_content(book, sheet, row, column) for column in range(left, left+width)])
    return rows


def verify_native_sources(export_path, root):
    """Сопоставляет каждый фактический модуль и события с текущими исходниками."""
    errors = []
    components = exported_components(export_path)
    files = [*root.glob('src/**/*.cls'), *root.glob('src/**/*.bas'), *root.glob('tests/**/*.cls'), *root.glob('tests/**/*.bas')]
    for path in files:
        native = components.get(path.stem)
        if native is None or normalized_vba(source_body(path)) != normalized_vba(native[1]):
            errors.append(f'Код книги отличается от исходника: {path.relative_to(root)}')
    documents = [(key, body) for key, (kind, body) in components.items() if kind == 100]
    active_documents = [body for _, body in documents if any(line != 'optionexplicit' and not line.startswith("'") for line in normalized_vba(body))]
    if len(documents) != 6 or len(active_documents) != 1:
        errors.append('Неожиданный состав событий книги и листов')
    events = normalized_vba((root/'tools/build_workbook/WorkbookHelpEvents.vba').read_text('utf-8-sig'))
    if not any(normalized_vba(body) == events for _, body in documents):
        errors.append('События книги отличаются от WorkbookHelpEvents.vba')
    if set(components) != {path.stem for path in files} | {key for key, _ in documents}:
        errors.append('Книга содержит отсутствующий в исходниках или временный VBA-компонент')
    return len(files), len(documents), errors


def verify_fresh(after_path, export_path, root):
    """Проверяет чистую сборку без старого расчёта, её настройки и реальный код."""
    book = read(after_path)
    sources, documents, errors = verify_native_sources(export_path, root)
    settings = {row[0][1]: row[1][1] for row in table(book, 'rngSystemSettings') if row[0][1]}
    defaults = {'Results.SnapshotStorage': 'Excel', 'Results.SnapshotName': 'NDM_Snapshot',
                'Mesh.StepMode': 'Auto 0.1', 'Mesh.BoundarySubdivisions': '2'}
    for key, expected in defaults.items():
        if str(settings.get(key)) != expected:
            errors.append(f'Начальная настройка отличается: {key}')
    cells = book['sheets']['Results']['cells']
    unexpected = [address for address, cell in cells.items() if
                  (cell.get('value') is not None or cell.get('formula') is not None)
                  and int(re.search(r'\d+$', address)[0]) not in (154, 155, 156)]
    if unexpected:
        errors.append(f'В чистой книге остались расчётные результаты: {unexpected[:8]}')
    order = ['rngNDMSectionGeometry', 'rngNDMSectionProperties', 'rngNDMMaterialDiagrams', 'rngNDMElementResults', 'rngNDMLoadCaseResults', 'rngNDMSectionContours', 'rngNDMSectionAnnotations']
    positions = [rectangle(book, name) for name in order]
    if len({item[2] for item in positions}) != 1 or any(a[3] >= b[1] for a, b in zip(positions, positions[1:])):
        errors.append('Нарушен горизонтальный порядок нижних таблиц')
    for name, sheet in book['sheets'].items():
        for address, cell in sheet['cells'].items():
            if '#REF!' in str(cell.get('formula') or '') or '#REF!' in str(cell.get('value') or ''):
                errors.append(f'Неверная ссылка в чистой книге: {name}!{address}')
    return dict(Status='FAIL' if errors else 'PASS', Mode='FreshBuild',
                WorkbookSHA256=hashlib.sha256(after_path.read_bytes()).hexdigest().upper(),
                SourceComponents=sources, DocumentComponents=documents,
                DefaultSettings=defaults, SavedResultsAbsent=not unexpected, LowerOrder=order, Errors=errors)


def verify(before_path, after_path, export_path, root):
    """Проверяет ровно разрешённые изменения структуры и опубликованный код."""
    old, new = read(before_path), read(after_path)
    errors, checks = [], 0
    for name in old['names']:
        try:
            sheet, *_ = rectangle(old, name)
        except ValueError:
            continue
        if name.startswith('_xlnm.') or sheet not in ('Config', 'Сечения'):
            continue
        a, b = table(old, name), table(new, name)
        if name == 'rngSystemSettings':
            # Вставка двух строк меняет адреса, а не уже введённые параметры.
            by_key = lambda rows: {row[0][1]: row[1] for row in rows if row[0][1] and '.' in row[0][1]}
            previous, actual = by_key(a), by_key(b)
            for key, value in previous.items():
                checks += 1
                if actual.get(key) != value:
                    errors.append(f'Изменён пользовательский параметр: {key}')
            added = set(actual) - set(previous)
            if added != {'Results.SnapshotStorage', 'Results.SnapshotName'}:
                errors.append(f'Неожиданные новые настройки: {sorted(added)}')
        else:
            checks += sum(map(len, a))
            if a != b:
                errors.append(f'Изменены значения/формулы входной таблицы: {name}')
    # Верхние таблицы сохраняются при структурном обновлении, включая формулы.
    lower = rectangle(old, 'rngNDMSectionGeometry')[2]
    for address, cell in old['sheets']['Results']['cells'].items():
        if int(re.search(r'\d+$', address)[0]) >= lower - 2:
            continue
        other = new['sheets']['Results']['cells'].get(address, {})
        checks += 1
        for field in ('value', 'formula', 'style'):
            if cell.get(field) != other.get(field):
                errors.append(f'Изменена прежняя верхняя таблица Results!{address}: {field}')
    added_names = set(new['names']) - set(old['names'])
    if added_names != {'rngNDMLoadCaseResults'}:
        errors.append(f'Неожиданные новые именованные диапазоны: {sorted(added_names)}')
    for name, width in [('rngNDMSectionGeometry', 16), ('rngNDMSectionContours', 17), ('rngNDMMaterialDiagrams', 11), ('rngNDMElementResults', 9), ('rngNDMSectionProperties', 6), ('rngNDMSectionAnnotations', 24)]:
        previous, actual = anchored_table(old, name, width), anchored_table(new, name, width)
        if name == 'rngNDMSectionProperties':
            previous = previous[:1] + [row for row in previous[1:] if row[1][1] == 'ALL' and row[2][1] not in ('LoadReferenceX', 'LoadReferenceY') and not str(row[2][1]).startswith('SP35.Slenderness')]
        elif name == 'rngNDMSectionAnnotations':
            previous = previous[:1] + [row for row in previous[1:] if not str(row[1][1]).startswith('CRACK_')]
        checks += sum(map(len, previous))
        if previous != actual:
            errors.append(f'Изменены прежние технические результаты: {name}')
    order = ['rngNDMSectionGeometry', 'rngNDMSectionProperties', 'rngNDMMaterialDiagrams', 'rngNDMElementResults', 'rngNDMLoadCaseResults', 'rngNDMSectionContours', 'rngNDMSectionAnnotations']
    positions = [rectangle(new, name) for name in order]
    if len({item[2] for item in positions}) != 1 or any(a[3] >= b[1] for a, b in zip(positions, positions[1:])):
        errors.append('Нарушен горизонтальный порядок нижних таблиц')
    sources, documents, source_errors = verify_native_sources(export_path, root)
    checks += sources
    errors.extend(source_errors)
    return dict(Status='FAIL' if errors else 'PASS', BeforeSHA256=hashlib.sha256(before_path.read_bytes()).hexdigest().upper(),
                WorkbookSHA256=hashlib.sha256(after_path.read_bytes()).hexdigest().upper(), Checks=checks,
                SourceComponents=sources, DocumentComponents=documents, VBENormalization='Identifier case, insignificant whitespace, native 15-significant-digit floating literals', AddedNames=sorted(added_names),
                LowerOrder=order, Errors=errors)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--before', type=Path)
    parser.add_argument('--fresh', action='store_true')
    parser.add_argument('--workbook', type=Path, required=True)
    parser.add_argument('--export', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    args = parser.parse_args()
    if not args.fresh and args.before is None:
        parser.error('--before is required unless --fresh is selected')
    root = Path(__file__).resolve().parents[2]
    result = verify_fresh(args.workbook, args.export, root) if args.fresh else verify(args.before, args.workbook, args.export, root)
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps(result, ensure_ascii=False))
    raise SystemExit(bool(result['Errors']))
