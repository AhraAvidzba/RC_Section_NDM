"""ДЛЯ ТЕСТОВ: полный первичный разбор уже закрытых записей Focused без пересчёта."""
from pathlib import Path
from datetime import datetime
import hashlib
import json
import re
import sys
import time


def TestReadJson(path):
    """Повторяет чтение одновременно обновляемого Run.json, не считая чтение остановкой."""
    for attempt in range(5):
        try:
            return json.loads(path.read_text(encoding='utf-8-sig'))
        except (OSError, ValueError):
            if attempt == 4:
                raise
            time.sleep(0.1)


def TestReviewFocusedPrimary():
    """Читает все записанные primary и отмечает факты; официальную полную приёмку не заменяет."""
    root = Path.cwd()
    directory = root / 'docs/regression/FullOnOff_2026-10-26_OOP_Run01/Focused'
    run = TestReadJson(directory / 'Run.json')
    records = []
    for item in run['Records']:
        record = {'ID': item['ID'], 'Kind': item['Kind'], 'Report': item['Report'],
                  'RecordPassed': item['Passed'], 'Errors': []}
        try:
            path = (root / item['Report']).resolve()
            assert path.is_relative_to(directory.resolve())
            raw = path.read_bytes()
            digest = hashlib.sha256(raw).hexdigest().upper()
            if item['Kind'] == 'VBA':
                assert digest == item['ReportSHA256']
            record['DigestRecordedInRun'] = 'ReportSHA256' in item
            text = raw.decode('utf-8-sig')
            record.update(SHA256=digest, Bytes=len(raw), LineCount=len(text.splitlines()))
            if item['Kind'] == 'VBA':
                errors = re.findall(r'^(?:FAIL[: ][^\r\n]*|RUNTIME ERROR[^\r\n]*|SCRIPT ERROR[^\r\n]*|WATCHDOG_FAILURE[^\r\n]*|RESTORE ERROR[^\r\n]*|CLEANUP ERROR[^\r\n]*)', text, re.M)
                errors += re.findall(r'^TOTAL[^\r\n]*failed=[1-9]\d*[^\r\n]*', text, re.M)
                owner = re.search(r'TEST_EXCEL_PROCESS: id=(\d+); startTicks=(\d+)', text)
                finished = re.findall(r'^SUITE_FINISHED: ([^;\r\n]+)', text, re.M)
                assert finished == [item['Macro']]
                assert 'WATCHDOG_COMPLETED: exit=0;' in text
                record.update(Errors=errors, FinishedMacros=finished,
                              OwnedExcel={'PID': int(owner[1]), 'StartTicks': owner[2]} if owner else None,
                              Totals=re.findall(r'^TOTAL[^\r\n]*', text, re.M),
                              FocusedResultIDs=re.findall(r'^FOCUSED_RESULT\|id=([^|\r\n]+)', text, re.M))
            else:
                native = json.loads(text)
                record['NativeRootSummary'] = {key: native[key] for key in ('Status', 'Passed', 'SourceSHA256') if key in native}
                if native.get('Passed') is False:
                    record['Errors'].append('Native report Passed=false')
            if not item['Passed']:
                record['Errors'].append('Focused Run record Passed=false')
        except Exception as error:
            record['Errors'].append(type(error).__name__ + ': ' + str(error))
        records.append(record)
    result = {'At': datetime.now().astimezone().isoformat(), 'Status': 'FOCUSED_PRIMARY_OBSERVATION_NOT_ACCEPTANCE',
              'SourceSHA256': run['SourceSHA256'], 'RunStatus': run['Status'], 'Records': records,
              'ReadRecords': len(records), 'ReportsWithErrors': [r['ID'] for r in records if r['Errors']], 'FullPass': False}
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    out = root / ('docs/regression/OOPLayers_2026-10-26_Run01/Observations/FocusedPrimaryReview_' + stamp + '.json')
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'Proof': str(out.relative_to(root)), 'ReadRecords': len(records),
                      'ReportsWithErrors': result['ReportsWithErrors'], 'LastRecord': records[-1]['ID'] if records else None}, ensure_ascii=True))
    return 1 if result['ReportsWithErrors'] else 0


if __name__ == '__main__':
    sys.exit(TestReviewFocusedPrimary())