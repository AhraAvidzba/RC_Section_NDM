"""ДЛЯ ТЕСТОВ: целиком читает адресные primary и публикует компактные доказательства принятой книги."""
import hashlib
import json
import re
import shutil
from pathlib import Path

P = Path(__file__).resolve().parent
R = Path('C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM')
E = R / 'docs/regression/MovablePlotContainer_2026-10-30_Run01'

def read(name):
    """ДЛЯ ТЕСТОВ: читает собственный JSON без потери кириллицы."""
    return json.loads((P / name).read_text(encoding='utf-8-sig'))

native, equality, integration, own, publication = [read(n) for n in [
    'Native02/Result.json', 'SourceEquality03.json', 'Integration07/Result.json', 'OwnClosures.json', 'Publication.json']]
assert all(v['Passed'] for v in [native, equality, integration, own, publication])
assert native['NamedPlaceholderPresentAtBuild'] and native['Components'] == 181
assert native['CatalogCells'] == 405 and native['Dropdowns'] == 106
assert equality['SourceComponents'] == 175 and equality['ExportComponents'] == 181
assert integration['PlacementStages'] == 6 and integration['AutomaticCalculationRecoveryNotice']
delta = read('CodeDelta03.json')
final = read('Final03/Result.json')
assert delta['Passed'] and delta['ProductionChangeOnlyEmptySnapshotMessage'] and delta['Other179BodiesEqual']
assert integration['SourceSHA256'] == delta['IntegrationSHA'] == integration['SourceAfterSHA256']
assert final['Passed'] and final['EmptySnapshotNotice'] and final['NoFalseSecondNotice']
assert final['BookSHA256'] == native['BookSHA256'] == final['SourceAfterSHA256'] == delta['FinalSHA']
assert all(v['IdenticalToPreviouslyViewed'] for v in read('FinalVisualEquivalence03.json'))
raw = []
for name, count in [
    ('modTestPlotConfig.RunPlotContainerTests', 24), ('modTestPlotConfig.RunAudit03AutoPlotTests', 86),
    ('modTestWorkbookInterface.RunFormationPreviewTests', 25), ('modTestWorkbookInterface.RunSectionWorkflowCalculationTests', 116)]:
    path = P / 'Directed04' / (name + '.txt')
    text = path.read_text(encoding='utf-8-sig')
    totals = re.findall(r'(?m)^TOTAL[^\r\n]*passed=(\d+); failed=(\d+)', text)
    assert totals and totals[-1] == (str(count), '0') and not re.search(r'(?m)^FAIL:', text), name
    raw.append({'Macro': name, 'Passed': count, 'Failed': 0, 'PrimarySHA256': hashlib.sha256(path.read_bytes()).hexdigest().upper()})
final_primary = (P / 'Final03/ContainerPrimary.txt').read_text(encoding='utf-8-sig')
assert 'passed=24; failed=0' in final_primary and not re.search(r'(?m)^FAIL:', final_primary)
for name in ['FormationMessage.txt', 'ManualRecoveryNotice.txt', 'AutomaticRecoveryMessage.txt']:
    assert 'автоматически создана' in (P / 'Integration07' / name).read_text(encoding='utf-8-sig')
assert (P / 'Integration07/ManualRecoveryNotice.txt').read_text(encoding='utf-8-sig').startswith('True|')
changed_inputs = []
for item in read('FinalBuildStart02.json')['SourceFence']:
    if hashlib.sha256((R / item['Path']).read_bytes()).hexdigest().upper() != item['SHA256']:
        changed_inputs.append(item['Path'])
assert set(changed_inputs) == {'src/Excel/modWorkbookCalculation.bas', 'tests/modTestPlotConfig.bas'}
for item in publication['Files']:
    assert hashlib.sha256((R / item['Target']).read_bytes()).hexdigest().upper() == item['SHA256']
assert read('Integration05/OwnTermination.json')['Exited']
assert read('Integration07/OwnTermination.json')['Exited'] and read('Integration07/OwnTermination.json')['UserExplicitlyApproved']
E.mkdir(exist_ok=False)
files = ['Backup.json', 'CatalogChanges.json', 'FinalBuildStart02.json', 'Build02.log',
         'Build/BuildOwnership.json', 'Build02/BuildOwnership.json', 'Native01/Result.json', 'Native01/Ownership.json',
         'Native01/CatalogReadback.json', 'Native01/NativeDropdowns.json', 'SourceEquality.json', 'OwnClosures.json',
         'Publication.json', 'VisualReview.json', 'TestNativeDefaults01.ps1', 'TestOwnClosures01.ps1',
         'TestVerifySource.ps1', 'TestPublish01.ps1', 'TestIntegration07.ps1', 'TestFinalize08.py',
         'Integration07/Result.json', 'Integration07/Ownership.json', 'Integration07/Placement.json',
         'Integration07/FormationMessage.txt', 'Integration07/ManualRecoveryNotice.txt', 'Integration07/CalculationMessage.txt',
         'Integration07/CalculationFileMessage.txt', 'Integration07/AutomaticRecoveryMessage.txt',
         'Integration07/modTestPlotConfig.RunPlotContainerTests.txt', 'Directed04/Macros.json',
         'Integration05/StoppedTestEvidence.json', 'Integration05/OwnTermination.json']
files += ['Native02/Result.json', 'Native02/Ownership.json', 'Native02/CatalogReadback.json', 'Native02/NativeDropdowns.json',
          'SourceEquality03.json', 'CodeDelta03.json', 'TestReviewCodeDelta03.ps1', 'TestFinalizeCode03.ps1',
          'Final03/Ownership.json', 'Final03/Result.json', 'Final03/EmptyRecoveryMessage.txt', 'Final03/ContainerPrimary.txt',
          'FinalVisualEquivalence03.json', 'TestNativeDefaults02.ps1', 'Integration07/OwnTermination.json',
          'Integration07/UserTerminationAuthorization.json']
files += ['Directed04/' + v['Macro'] + '.txt' for v in raw]
files += [f'Directed{i:02}/modTestPlotConfig.RunPlotContainerTests.txt' for i in [1, 2, 3]]
manifest = []
for name in files:
    dest = E / name
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(P / name, dest)
    manifest.append({'Path': name, 'SHA256': hashlib.sha256(dest.read_bytes()).hexdigest().upper()})
proof = {'Status': 'DIRECTED_CONTAINER_NATIVE_REOPEN_AND_PUBLICATION_PASS', 'PrimaryLogsEntirelyRead': True,
         'Directed': raw, 'DirectedPassed': 251, 'DirectedFailed': 0, 'FinalBuildContainerChecks': 24,
         'PlacementStages': read('Integration07/Placement.json'), 'FormationManualAndAutomaticRecoveryNoticeVerified': True,
         'Native181AndSourceEquality': True, 'FinalBookSHA256': native['BookSHA256'],
         'HistoricalStoppedOwnTestForceTerminated': read('Integration05/OwnTermination.json'),
         'UserApprovedFinalTestTermination': read('Integration07/OwnTermination.json'),
         'PostIntegrationChange': delta, 'EmptySnapshotRecoveryNotice': True,
         'AllOwnTerminal': True, 'FullAcceptanceClaim': False, 'Manifest': manifest}
(E / 'Processed.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2), encoding='utf-8')
report = f'''# Переносимый контейнер схемы — 30.10.2026

Сборщик создаёт пустой `chtNDMSectionPlot` по `WorkbookLayout.json`. Отрисовка и очистка находят единственный контейнер по всем обычным листам текущей книги, сохраняя выбранный лист, положение и размеры. Дубликаты отклоняются до изменения обоих рисунков, с именем контейнера и названиями листов. Посторонние диаграммы сохраняются.

Если контейнера нет, создаётся новый на листе «Расчет» с явным предупреждением. Оно возвращается через итог формирования, ручного обновления и автоматического вывода после расчёта, включая отключённые обычные уведомления. При отсутствии и контейнера, и листа «Расчет» выдаётся понятная причина. Отдельные листы диаграмм не входят в контракт: переносить нужно встроенную диаграмму на обычный лист, сохраняя имя.

Справка Plot.Enabled и комментарии сборки обновлены. Основная книга пересобрана с нуля через Workbooks.Add и компиляцию VBA578; ручные изменения не переносились по прямому поручению пользователя, полный резерв сохранён. Значения PointedSection/HollowRectangle из предыдущего поручения сохранены в каталоге и подтверждены нативно.

После интеграции Build02 выполнена точечная финализация VBA на этой чистой основе: только сообщение ошибки пустого снимка дополнено текстом восстановления области, добавлен тестовый перехват этой ошибки. CodeDelta03 сравнил все181 нативных компонента:179 совпали, производственное изменение ровно одно ожидаемое выражение сообщения, второй изменённый модуль тестовый. На Final03 повторно приняты24 проверки контейнера, создание без данных с честным сообщением и отсутствие ложного повторного сообщения; native181/equality и все405/106 также проверены на Final03. Интеграционный результат Build02 не приписывается новой SHA без этой проверки узкой дельты.

Приёмка: четыре адресных набора 251/0, все raw primary целиком прочитаны. На окончательной чистой сборке повторены 24 проверки контейнера. Нативная интеграция проверила сохранение/повторное открытие, preview, ручное обновление Excel, автоматический вывод после расчёта, формирование File и чтение File. Все шесть этапов сохранили лист, координаты 73/91 и размеры 720×460 без появления второй схемы на «Расчет». Восстановление удалённого контейнера с сообщением отдельно проверено при формировании, ручном обновлении и расчёте. Расчёты выполнялись на малом круглом сечении 88 бетонных элементов/8 стержней/одно сочетание. Полный набор ON/OFF и Focused не запускался, новый общий PASS не заявляется.

Нативная сохранённая книга: 181 компонент, равенство 175 исходных и 6 документных модулей, 405 полей каталога, 106 списков, рамки/заливки/порядок листов, готовый именованный контейнер и связанные абзацы справки. Основная книга содержит пустые нагрузки, Results и схему; расчёты проверок остались в отдельных копиях.

Промежуточные отказы сохранены: текстовый формат RGB в тестовом шаблоне, COM-обёртка вместо проверки постоянной метки, Clipboard.Paste, неполный синтетический fixture и синтаксис PowerShell ChartObjects(). Интеграция05 ошибочно переключила File до создания снимка; штатный отказ отсутствующего файла оставил собственный Excel12676/639289690301059923 в VBA break-mode. После сохранения доказательств, попытки UI-завершения и проверки идентичности этот остановленный тест был принудительно завершён. Исправленная последовательность File и все окончательные проверки прошли. Собственный Excel18188/639289698228597295 после успешной Integration07 и Quit остался активным; повторный Quit отклонён auto-review, затем пользователь явно разрешил завершить только PID18188. Он завершён после повторной сверки идентичности и сохранения разрешения. Все15 собственных процессов терминальны; два принудительных завершения не скрываются как normalQuit.

Книга SHA256 `{native['BookSHA256']}`. Доказательства — `docs/regression/MovablePlotContainer_2026-10-30_Run01/Processed.json`. Ветка `codex/movable-plot-container`; main не изменялась. Пользовательские отчёты, outputs и отдельная книга 30 сочетаний не затрагивались.
'''
(R / 'docs/reports/MovablePlotContainer_2026-10-30.md').write_text(report, encoding='utf-8')
progress = R / 'docs/geometry_refactor/PROGRESS.md'
text = progress.read_text(encoding='utf-8-sig')
first, rest = text.split('\n', 1)
entry = '\n## 30.10.2026 — Переносимый именованный контейнер схемы\n\nЕдинственный chtNDMSectionPlot ищется по всем обычным листам; перемещение и размеры сохраняются, дубликаты отклоняются до очистки. Отсутствующий контейнер восстанавливается на «Расчет» с предупреждением в итогах формирования, ручного обновления и расчёта, включая ошибку пустого снимка. Чистая сборка, адресные251/0 и финальные24, сохранение/повторное открытие и Excel/File приняты, VBA181/equality/узкая дельта и все15 own проверены. Два тестовых процесса завершены принудительно после proofs; второй — по отдельному явному разрешению пользователя. [Отчёт](../reports/MovablePlotContainer_2026-10-30.md).\n'
progress.write_text(first + '\n' + entry + rest, encoding='utf-8')
print(json.dumps({k: v for k, v in proof.items() if k not in ['Manifest', 'PlacementStages', 'Directed']}, ensure_ascii=False))
