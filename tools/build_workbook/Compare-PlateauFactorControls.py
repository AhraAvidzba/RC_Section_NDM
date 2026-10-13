"""Сравнивает фактические статусы прежней книги и нового default в обоих режимах."""
import argparse
import hashlib
import json
import re
from pathlib import Path


def read(path):
    text = path.read_text(encoding='utf-8-sig')
    if 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' not in text:
        raise ValueError('Run did not complete successfully: '+path.name)
    if re.search(r'^FAIL:|RUNTIME ERROR|SCRIPT ERROR|RESTORE ERROR', text, re.M):
        raise ValueError('Assertion or runtime failure: '+path.name)
    rows = {}
    for match in re.finditer(r'^FOCUSED_RESULT\|([^\r\n]+)', text, re.M):
        record = dict(item.split('=',1) for item in match[1].split('|'))
        if record['id'] in rows:
            raise ValueError('Duplicate case: '+record['id'])
        rows[record['id']] = record
    if len(rows) != 12 or not re.search(r'TOTAL_FACTOR_CONTROLS: factor=0\.(?:0001|001); cases=12', text):
        raise ValueError('Incomplete 12-case comparison: '+path.name)
    unresolved = set()
    for match in re.finditer(r'^FOCUSED_UNRESOLVED\|([^\r\n]+)', text, re.M):
        record = dict(item.split('=',1) for item in match[1].split('|'))
        unresolved.add((record['id'], record['kind']))
    return rows, unresolved, hashlib.sha256(path.read_bytes()).hexdigest()


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--directory', type=Path, required=True)
    args = parser.parse_args()
    errors, records, hashes, ok = [], [], {}, 0
    for mode in ('No','Yes'):
        before, old_unresolved, digest = read(args.directory/('Baseline_'+mode+'.txt'))
        hashes['Baseline_'+mode+'.txt'] = digest
        after, new_unresolved, digest = read(args.directory/('Release_'+mode+'.txt'))
        hashes['Release_'+mode+'.txt'] = digest
        if set(before) != set(after):
            errors.append(mode+': case inventory changed')
        if new_unresolved-old_unresolved:
            errors.append(mode+': new unresolved result kinds '+str(sorted(new_unresolved-old_unresolved)))
        for name, old in before.items():
            new = after[name]
            equal = old == new
            if not equal:
                errors.append(mode+'/'+name+': statuses changed')
            if old['overall'] == 'OK':
                ok += 1
            records.append({'mode':mode, 'id':name, 'before':old, 'after':new, 'statusesEqual':equal})
    report = {'AllPassed': not errors, 'ComparedPairs':len(records), 'NativeCalculations':len(records)*2,
              'PreviouslyOKPairs':ok, 'NewNumericalFailures':0 if not errors else None,
              'ReportSHA256':hashes, 'Records':records, 'Errors':errors}
    (args.directory/'StatusComparison.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k:v for k,v in report.items() if k not in ('Records','ReportSHA256')}, ensure_ascii=False))
    raise SystemExit(0 if report['AllPassed'] else 1)
