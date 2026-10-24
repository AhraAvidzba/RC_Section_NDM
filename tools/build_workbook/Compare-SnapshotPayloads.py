"""ДЛЯ ТЕСТОВ: независимый разбор NDMSNP04 и сравнение всех матриц двух расчётов."""
from __future__ import annotations

import argparse
import hashlib
import json
import struct
from pathlib import Path

MAGIC = b"NDMSNP04"


class TestSnapshotReader:
    """Проверяет границы и SHA256; не использует VBA или его двоичный читатель."""

    def __init__(self, path: Path):
        self.data = path.read_bytes()
        self.position = 0
        self.limit = len(self.data) - 72
        if self.limit < 32 or self.data[-8:] != MAGIC:
            raise ValueError("Truncated snapshot")
        if hashlib.sha256(self.data[:self.limit]).hexdigest().upper().encode("ascii") != self.data[-72:-8]:
            raise ValueError("Snapshot SHA256 is invalid")

    def TestTake(self, count: int) -> bytes:
        """Читает ровно заданный блок, отклоняя отрицательные длины и выход за footer."""
        if count < 0 or count > self.limit - self.position:
            raise ValueError("Invalid snapshot length")
        block = self.data[self.position:self.position + count]
        self.position += count
        return block

    def TestLong(self) -> int:
        """Декодирует штатный little-endian Long без десятичного преобразования."""
        return struct.unpack("<i", self.TestTake(4))[0]

    def TestText(self) -> bytes:
        """Сохраняет точные UTF-16LE байты, включая пустую строку как отдельный тип."""
        count = self.TestLong()
        if count < 0 or count > 131072 or count % 2:
            raise ValueError("Invalid UTF-16 length")
        return self.TestTake(count)

    def TestScalar(self) -> tuple:
        """Оставляет исходные биты чисел; Boolean сравнивается как логический тип."""
        kind = self.TestTake(1)[0]
        if kind == 0:
            return 0, b""
        if kind == 1:
            return 1, self.TestText()
        if kind == 2:
            return 2, self.TestTake(8)
        if kind == 3:
            return 3, struct.unpack("<d", self.TestTake(8))[0] != 0
        raise ValueError("Invalid scalar type")

    def TestMatrix(self) -> tuple:
        """Декодирует скалярные, плотные и серийные колонки в их исходном порядке."""
        rows, columns = self.TestLong(), self.TestLong()
        if not 1 <= rows <= 10000000 or not 1 <= columns <= 256 or rows * columns > 100000000:
            raise ValueError("Invalid matrix dimensions")
        result = []
        for _ in range(columns):
            mask = self.TestTake(1)[0]
            if mask == 5:
                values = [self.TestScalar() for _ in range(rows)]
            elif mask == 4:
                values = []
                count = self.TestLong()
                if not 1 <= count <= rows:
                    raise ValueError("Invalid RLE count")
                for _ in range(count):
                    length = self.TestLong()
                    value = self.TestScalar()
                    if length < 1 or length > rows - len(values):
                        raise ValueError("Invalid RLE length")
                    values.extend([value] * length)
                if len(values) != rows:
                    raise ValueError("Incomplete RLE column")
            elif 0 <= mask <= 3:
                flags = self.TestTake(rows)
                numbers = self.TestTake(rows * 8) if mask & 1 else b""
                codes, strings = (), ()
                if mask & 2:
                    codes = struct.unpack(f"<{rows}i", self.TestTake(rows * 4))
                    count = self.TestLong()
                    if not 1 <= count <= rows:
                        raise ValueError("Invalid dictionary size")
                    strings = [self.TestText() for _ in range(count)]
                values = []
                for index, flag in enumerate(flags):
                    if flag == 0:
                        value = (0, b"")
                    elif flag == 1 and mask & 2 and 1 <= codes[index] <= len(strings):
                        value = (1, strings[codes[index] - 1])
                    elif flag in (2, 3) and mask & 1:
                        number = numbers[index * 8:index * 8 + 8]
                        value = (2, number) if flag == 2 else (3, struct.unpack("<d", number)[0] != 0)
                    else:
                        raise ValueError("Invalid dense value")
                    values.append(value)
            else:
                raise ValueError("Invalid column type")
            result.append(values)
        return rows, result

    def TestRead(self) -> tuple:
        """Требует полный законченный контейнер без дубликатов ключей и хвостовых байтов."""
        if self.TestTake(8) != MAGIC or self.TestLong() != 4:
            raise ValueError("Unsupported snapshot format")
        run_id, elapsed = self.TestTake(8), self.TestTake(8)
        count = self.TestLong()
        if not 1 <= count <= 64:
            raise ValueError("Invalid table count")
        tables = {}
        for _ in range(count):
            key = self.TestText().decode("utf-16-le", "surrogatepass")
            if key in tables:
                raise ValueError("Duplicate snapshot table")
            tables[key] = self.TestMatrix()
        if self.position != self.limit:
            raise ValueError("Unexpected snapshot tail")
        return run_id, elapsed, tables


def normalized(value: tuple) -> tuple:
    """Excel может менять только знак точного нуля; прочие числа сравниваются побитно."""
    if value[0] == 2 and value[1] in (b"\0" * 8, b"\0" * 7 + b"\x80"):
        return 2, b"\0" * 8
    return value


def main() -> None:
    """При разных запусках исключает только согласованный RunID и измеренное время."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline", type=Path, required=True)
    parser.add_argument("--candidate", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    run_a, _, a = TestSnapshotReader(args.baseline).TestRead()
    run_b, _, b = TestSnapshotReader(args.candidate).TestRead()
    if a.keys() != b.keys() or len(a) != 16:
        raise ValueError("The complete sixteen-table contracts differ")
    checks, rows = 0, []
    for key, (height, columns) in a.items():
        other_height, other = b[key]
        if height != other_height or len(columns) != len(other):
            raise ValueError("Matrix dimensions differ: " + key)
        for column, (values, following) in enumerate(zip(columns, other)):
            for row, (x, y) in enumerate(zip(values, following)):
                checks += 1
                if (key.startswith("rngNDM") or key.startswith("LoadCase.")) and column == 0 and row > 0:
                    if x != (2, run_a) or y != (2, run_b):
                        raise ValueError("Inconsistent RunID: " + key)
                    continue
                if key == "Summary" and column == 0 and row == 3:
                    if x[0] != 2 or y[0] != 2:
                        raise ValueError("Invalid elapsed time")
                    continue
                if normalized(x) != normalized(y):
                    raise ValueError(f"Value changed: {key}[{row + 1},{column + 1}]")
        rows.append(dict(Key=key, Rows=height, Columns=len(columns), AllValuesEqual=True))
    proof = dict(Status="PASS", Tables=rows, CellsCompared=checks,
                 BaselineSHA256=hashlib.sha256(args.baseline.read_bytes()).hexdigest().upper(),
                 CandidateSHA256=hashlib.sha256(args.candidate.read_bytes()).hexdigest().upper(),
                 RunIDsInternallyConsistent=True, Exclusions=["RunID between separate calculations", "Elapsed time"],
                 NumberEquivalence="Only the sign of exact zero; every other Double bit is unchanged")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(proof, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(dict(Status="PASS", Tables=len(rows), CellsCompared=checks)))


if __name__ == "__main__":
    main()
