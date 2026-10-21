"""Summarize only completed measured samples; warmups remain in raw evidence."""
import csv
import json
import statistics
from pathlib import Path

HERE = Path(__file__).resolve().parent
summary = {}
for path in HERE.glob('*.csv'):
    rows = list(csv.DictReader(path.read_text(encoding='utf-8-sig').splitlines()))
    if not rows or 'variant' not in rows[0]:
        continue
    receipt = path.with_suffix('.txt')
    receipt_text = receipt.read_text(encoding='utf-8-sig') if receipt.exists() else ''
    groups = {}
    for row in rows:
        if int(row['repetition']) == 0:
            continue
        groups.setdefault(row['variant'], []).append(row)
    result = {}
    for name, samples in groups.items():
        item = {'samples': len(samples)}
        for field in samples[0]:
            if field in ['variant', 'repetition']:
                continue
            try:
                values = [float(row[field]) for row in samples]
                item[field] = statistics.median(values)
                if field in ['formation', 'calculation']:
                    item[field + 'Range'] = [min(values), max(values)]
                if field == 'mismatch':
                    item['totalMismatches'] = sum(values)
            except ValueError:
                item[field] = sorted(set(row[field] for row in samples))
        result[name] = item
    if 'Full' in result:
        for name, item in result.items():
            for field in ['formation', 'calculation']:
                item[field + 'Speedup'] = result['Full'][field] / item[field]
            item['total'] = item['formation'] + item['calculation']
            item['totalSpeedup'] = (result['Full']['formation'] + result['Full']['calculation']) / item['total']
    summary[path.stem] = {'variants': result, 'rowsIncludingWarmups': len(rows), 'completed': 'EXPERIMENT_COMPLETE:' in receipt_text, 'failuresReported': 'FAIL:' in receipt_text}
    print(path.stem)
    for name, item in result.items():
        print(f"  {name:12} {item['samples']}x form={item['formation']:.3f}s ({item.get('formationSpeedup',0):.2f}x) calc={item['calculation']:.3f}s ({item.get('calculationSpeedup',0):.2f}x) diff={item.get('totalMismatches', 'intentional')}")
(HERE / 'Summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
