"""Independently validate JSON emitted by Excel and decode its exact Double tags."""
from pathlib import Path
import json
import math
import struct

HERE = Path(__file__).resolve().parent

def decode(value):
    if isinstance(value, str):
        if value.startswith('~~'):
            return value[1:]
        if len(value) == 18 and value.startswith('~D'):
            result = struct.unpack('>d', bytes.fromhex(value[2:]))[0]
            assert math.isfinite(result)
            return result
    return value

records = []
for path in sorted(HERE.glob('*_JsonExact_rng*.json')):
    rows, cols, descriptors = json.loads(path.read_text(encoding='utf-8-sig'))
    assert len(descriptors) == cols
    data = []
    for header, flag, values in descriptors:
        assert isinstance(header, str)
        assert flag in [0, 1]
        values = [values] * (rows - 1) if flag == 0 else values
        assert len(values) == rows - 1
        data.append([header] + [decode(value) for value in values])
    other = path.with_name(path.name.replace('_JsonExact_', '_JsonExactNoPlot_'))
    other_rows, other_cols, other_descriptors = json.loads(other.read_text(encoding='utf-8-sig'))
    assert (rows, cols) == (other_rows, other_cols)
    compared = 0
    for index, (header, flag, values) in enumerate(other_descriptors):
        assert header == data[index][0]
        if header == 'RunID':
            continue
        values = [values] * (rows - 1) if flag == 0 else values
        assert [decode(value) for value in values] == data[index][1:]
        compared += rows - 1
    records.append({'file': path.name, 'rows': rows, 'columns': cols,
                    'characters': len(path.read_text(encoding='utf-8-sig')),
                    'normalTableCells': rows * cols, 'sameCellsWithAndWithoutPlot': compared})
result = {'passed': bool(records), 'files': records,
          'scope': 'JSON syntax, dimensions, exact-double decoding, equality of stored auxiliary data with/without plot except RunID; not production acceptance'}
(HERE / 'JsonVerification.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(result, ensure_ascii=False, indent=2))
