"""Record isolated paired workflow times and numerical fingerprints."""
import argparse
import hashlib
import json
import re
import statistics
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory', type=Path, required=True)
    args = parser.parse_args()
    root = args.directory
    errors, records, comparisons, files = [], {}, [], []

    def load(name, kind):
        path = root/name
        if not path.is_file():
            errors.append('Missing: ' + name)
            return ''
        text = path.read_text(encoding='utf-8-sig')
        expected = hashlib.sha256((root/('AuxiliaryModeBaseline.xlsm' if kind == 'Baseline' else 'AuxiliaryModeFinal.xlsm')).read_bytes()).hexdigest()
        source = re.search(r'(?m)^SOURCE_SHA256: ([0-9A-Fa-f]{64})$', text)
        if not source or source.group(1).lower() != expected or 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' not in text or re.search(r'(?m)^FAIL:|RUNTIME ERROR', text):
            errors.append('Unsuccessful or changed input: ' + name)
        files.append({'path': name, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()})
        return text

    def metrics(line):
        result = {}
        for key, value in re.findall(r'(?:^|; )([A-Za-z0-9]+)=([^;\r\n]+)', line):
            try:
                result[key] = float(value.replace(',', '.'))
            except ValueError:
                result[key] = value
        return result

    def samples(text, prefix):
        # A VBA fingerprint contains the complete multiline Results tables.
        # Comparing only the first metrics line would omit the numeric data.
        boundaries = list(re.finditer(r'(?m)^(?:FAMILY_WORKFLOW|PASSIVE_FAMILY|FAMILY_STAGES|MIXED_WORKFLOW|COM_WORKBOOK_AFTER_SUITE|MODE_AFTER_SUITE|SUITE_FINISHED):', text))
        values = []
        for index, match in enumerate(boundaries):
            if not text[match.start():].startswith(prefix + ':'):
                continue
            end = boundaries[index+1].start() if index+1 < len(boundaries) else len(text)
            block = text[match.start():end].strip()
            first_line = block.splitlines()[0].split(': ', 1)[1]
            if prefix != 'MIXED_WORKFLOW':
                name, first_line = first_line.split('; ', 1)
            else:
                name = 'mixed'
            item = metrics(first_line)
            fingerprint = block.split('fingerprint=', 1)[1].strip() if 'fingerprint=' in block else ''
            item['fingerprint'] = hashlib.sha256(fingerprint.encode('utf-8')).hexdigest() if fingerprint else None
            item['fingerprintLines'] = len(fingerprint.splitlines())
            if item['fingerprintLines'] < 100:
                errors.append('Incomplete multiline Results fingerprint: ' + prefix + '/' + name)
            values.append((name, item))
        return values

    for mode in ('Off', 'On'):
        for kind in ('Baseline', 'Current'):
            workflow = load(f'FinalWorkflow_{mode}_{kind}.txt', kind)
            rows = dict(samples(workflow, 'FAMILY_WORKFLOW'))
            if len(rows) != 9 or len(re.findall(r'(?m)^PASSIVE_FAMILY: .*addedSolves=0;', workflow)) != 9:
                errors.append('Expected nine complete workflows and passive plots: ' + mode + '/' + kind)
            staged = load(f'FinalStagesMixed_{mode}_{kind}.txt', kind)
            stages = [item for _, item in samples(staged, 'FAMILY_STAGES')]
            mixed = [item for _, item in samples(staged, 'MIXED_WORKFLOW')]
            if len(stages) != 3 or len(mixed) != 1:
                errors.append('Expected three staged samples and one mixed workflow: ' + mode + '/' + kind)
            for item in list(rows.values()) + stages + mixed:
                if item.get('diagramExtension') != ('No' if mode == 'Off' else 'Yes'):
                    errors.append('Actual diagram mode mismatch: ' + mode + '/' + kind)
                if not isinstance(item.get('seconds'), (float, int)) or item['seconds'] <= 0:
                    errors.append('Invalid actual workflow time: ' + mode + '/' + kind)
            records[mode + '/' + kind] = {'families': rows, 'stages': stages, 'mixed': mixed,
                'warmMedianSeconds': statistics.median(x['seconds'] for x in stages[1:]) if len(stages) == 3 else None}
        before, after = records[mode + '/Baseline'], records[mode + '/Current']
        if set(before['families']) != set(after['families']):
            errors.append('Different measured family sets: ' + mode)
        for name in sorted(set(before['families']) & set(after['families'])):
            a, b = before['families'][name], after['families'][name]
            fingerprint_same = bool(a.get('fingerprint')) and a.get('fingerprint') == b.get('fingerprint')
            if not fingerprint_same:
                errors.append('Workflow fingerprint differs: ' + mode + '/' + name)
            comparisons.append({'mode': mode, 'family': name, 'baselineSeconds': a.get('seconds'),
                'currentSeconds': b.get('seconds'), 'deltaSeconds': b.get('seconds', 0)-a.get('seconds', 0),
                'fingerprintSame': fingerprint_same, 'under10Seconds': b.get('seconds', 1000) <= 10})
        for label, a_rows, b_rows in (('staged', before['stages'], after['stages']), ('mixed', before['mixed'], after['mixed'])):
            if len(a_rows) != len(b_rows):
                errors.append('Different measurement sample counts: ' + mode + '/' + label)
            for index, (a, b) in enumerate(zip(a_rows, b_rows), 1):
                if not a.get('fingerprint') or a.get('fingerprint') != b.get('fingerprint'):
                    errors.append('Measurement fingerprint differs: ' + mode + '/' + label + str(index))
    result = {'status': 'PASS' if not errors else 'INCOMPLETE_OR_FAILED', 'modes': ['Off', 'On'],
        'records': records, 'comparisons': comparisons, 'artifacts': files, 'errors': errors,
        'under10Seconds': all(x['under10Seconds'] for x in comparisons) and len(comparisons) == 18,
        'scope': 'Nine full single-LC workflows per mode with preparation/report/Results/plot; paired staged and mixed HollowOffset',
        'measurementCondition': 'Run only after main/native jobs; no concurrent heavy regression',
        'notUsedAsBenchmark': 'Aggregate suite wall times and COM/dialog retry waits'}
    (root/'FinalPerformanceEvidence.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps({'status': result['status'], 'comparisons': len(comparisons), 'errors': errors}))
    return 0 if not errors else 1


if __name__ == '__main__':
    raise SystemExit(main())
