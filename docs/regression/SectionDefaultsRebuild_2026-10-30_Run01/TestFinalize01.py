"""ДЛЯ ТЕСТОВ: фиксирует результаты фактической пересборки, нативного чтения и публикации без расширения приёмки."""
import hashlib
import json
import shutil
from pathlib import Path

P = Path(__file__).resolve().parent
R = Path('C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM')
E = R / 'docs/regression/SectionDefaultsRebuild_2026-10-30_Run01'
E.mkdir(exist_ok=False)

def read(name):
    """ДЛЯ ТЕСТОВ: читает собственный UTF-8 JSON с необязательным BOM."""
    return json.loads((P / name).read_text(encoding='utf-8-sig'))

native, equality, own, published = [read(n) for n in ['Native01/Result.json', 'SourceEquality.json', 'OwnClosures.json', 'Publication.json']]
assert all(x['Passed'] for x in [native, equality, own, published])
assert native['Components'] == 181 and native['CatalogCells'] == 405 and native['Dropdowns'] == 106
assert equality['SourceComponents'] == 175 and equality['ExportComponents'] == 181
assert len(own['OwnProcesses']) == 2
for item in published['Files']:
    assert hashlib.sha256((R / item['Target']).read_bytes()).hexdigest().upper() == item['SHA256']
    assert hashlib.sha256(Path(item['Backup']).read_bytes()).hexdigest().upper() == item['PreviousSHA256']
visual = {'Reviewed': True, 'Views': ['PointedSection', 'HollowRectangle', 'Config', 'Help', 'Calculation', 'Results'],
          'ReviewedInsideConversation': True, 'ClaimedDeliveredToUser': False,
          'RequestedValuesVisible': True, 'ExistingTableStyleAndBoundsRetained': True}
(P / 'VisualReview.json').write_text(json.dumps(visual, ensure_ascii=False, indent=2), encoding='utf-8')
names = ['Backup.json', 'RequestedDefaults.json', 'CatalogChanges.json', 'CleanBuildStart.json',
         'Build.log', 'Build/BuildOwnership.json', 'Native01/Result.json', 'Native01/Ownership.json',
         'Native01/CatalogReadback.json', 'Native01/NativeDropdowns.json', 'SourceEquality.json',
         'OwnClosures.json', 'Publication.json', 'VisualReview.json', 'TestSetDefaults01.py',
         'TestNativeDefaults01.ps1', 'TestOwnClosures01.ps1', 'TestVerifySource.ps1', 'TestPublish01.ps1', 'TestFinalize01.py']
manifest = []
for name in names:
    target = E / name
    target.parent.mkdir(exist_ok=True, parents=True)
    shutil.copyfile(P / name, target)
    manifest.append({'Path': name, 'SHA256': hashlib.sha256(target.read_bytes()).hexdigest().upper(), 'Bytes': target.stat().st_size})
processed = {'Status': 'CLEAN_REBUILD_REQUESTED_DEFAULTS_NATIVE_VERIFIED_PUBLISHED',
             'ChangedBuildDefaults': 10, 'CatalogValuesReadBack': 405, 'NativeComponents': 181,
             'SourceEquality': True, 'NativeDropdowns': 106, 'OwnProcessesClosed': 2,
             'UserMainBackupSHA256': read('Backup.json')['BookSHA256'],
             'BookSHA256': published['Files'][0]['SHA256'], 'AllCodeSHA256': published['Files'][1]['SHA256'],
             'CleanLoadsResultsAndPlot': True, 'FullAcceptanceClaim': False,
             'Branch': 'codex/section-defaults-rebuild', 'Manifest': manifest}
(E / 'Processed.json').write_text(json.dumps(processed, ensure_ascii=False, indent=2), encoding='utf-8')
report = '''# Чистая пересборка и значения сечений — 30.10.2026

Основная книга пересобрана из актуальных исходников через `Workbooks.Add`, с компиляцией VBA (команда 578). Каталог `tools/build_workbook/GeometryTables.json` задаёт значения для каждой последующей чистой сборки. Изменены ровно 10 значений каталога; остальные типы, комментарии и раскладка не менялись.

| Параметры | PointedSection | HollowRectangle |
|---|---|---|
| Габариты | B=800, H=250 | B=500, H=800, R=180 |
| Боковые параметры / отверстие | Left.W=0, Left.R1=100, Left.R2=0; Right.W=150, Right.R1=40, Right.R2=50 | Bo=200, Ho=500, Ro=30, Xo=Yo=0 |
| Основной ряд сверху и снизу | as=40, d=20, n=8, s0=s1=0, IncludeBoth | as=40, d=20, n=5, s0=s1=0, ExcludeBoth |
| Основной ряд слева и справа | as=40, d=20, n=3, s0=s1=0, ExcludeBoth | as=40, d=20, n=6, s0=s1=0, IncludeBoth |
| Ряды отверстия | — | as=40, d=20, s0=s1=0; n и endMode — проекционные прочерки |

Во всех дополнительных рядах обоих типов d2/d3 пустые, a2/a3=0, bind2/bind3=EachBar, pos2/pos3=Stacked. Значения сверены со скриншотами пользователя и затем прочитаны из сохранённой, повторно открытой книги Excel.

Нативная проверка: 405 полей каталога совпали; 106 выпадающих списков принимают отдельные варианты и отклоняют склейки через запятую, точку с запятой и постороннее значение. Исходные значения после проверки восстановлены без сохранения. Проверены границы таблиц, заливки, рамки, выравнивание комментариев, порядок листов, пустые нагрузки/Results/схема. Просмотрены два изменённых блока и представления всех пяти листов. Все 175 исходных и 6 документных VBA-компонентов соответствуют источникам. Оба собственных Excel вышли нормально, без принудительного завершения.

Основной файл заменён чистой сборкой; прежняя пользовательская книга полностью сохранена в резерве. Настоящая книга 30 сочетаний, чужие отчёты и outputs не менялись. Исполняемый VBA в рамках этой задачи не изменялся, расчётные тесты и новый полный набор не запускались; новый общий PASS не заявляется.

Книга SHA256: `{book}`. AllCode SHA256: `{code}`.

Доказательства: `docs/regression/SectionDefaultsRebuild_2026-10-30_Run01/Processed.json`, нативные чтения, SourceEquality, OwnClosures и Publication. Ветка `codex/section-defaults-rebuild`; слияние в main выполняется после согласования результата.
'''.format(book=processed['BookSHA256'], code=processed['AllCodeSHA256'])
(R / 'docs/reports/SectionDefaultsRebuild_2026-10-30.md').write_text(report, encoding='utf-8')
progress = R / 'docs/geometry_refactor/PROGRESS.md'
text = progress.read_text(encoding='utf-8-sig')
first, rest = text.split('\n', 1)
entry = '''
## 30.10.2026 — Чистая пересборка со значениями двух сечений

Значения PointedSection и HollowRectangle со скриншотов закреплены в каталоге сборки: 10 изменений, остальные параметры сохранены. Основная книга пересобрана с нуля; резерв пользовательской книги сохранён. Нативные 181/equality, 405 полей каталога, 106 списков, оформление, пустые нагрузки/Results/схема и два собственных выхода приняты. Расчёты и полный набор не запускались. [Отчёт](../reports/SectionDefaultsRebuild_2026-10-30.md).
'''
progress.write_text(first + '\n' + entry + rest, encoding='utf-8')
print(json.dumps({k: v for k, v in processed.items() if k != 'Manifest'}, ensure_ascii=False))
