"""ДЛЯ ТЕСТОВ: сравнивает чистый IO и побайтовую совместимость двух кодеков."""
from __future__ import annotations

import argparse
import hashlib
import json
import statistics
from pathlib import Path


def summarize(directory: Path, file_only: bool = False) -> dict:
    """Требует пять точных круговых проверок каждого действия вне прогрева."""
    measurements = json.loads((directory / "Measurements.json").read_text("utf-8-sig"))
    groups: dict[tuple, list[float]] = {}
    for row in measurements:
        if file_only and row["Action"] not in ("SaveFile", "LoadFile"):
            continue
        if not row["ExactAllCells"] or not row["Synthetic"]:
            raise ValueError("Codec round trip was not exact")
        if not row["Warmup"]:
            groups.setdefault((row["Elements"], row["States"], row["Action"]), []).append(row["Seconds"])
    if any(len(rows) < 5 for rows in groups.values()):
        raise ValueError("Five observations per action are required")
    return {key: {"Median": statistics.median(rows), "Min": min(rows), "Max": max(rows), "Observations": len(rows)}
            for key, rows in groups.items()}


def main() -> None:
    """Сверяет сохранённые файлы независимо от читателя VBA и сводит медианы."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline", type=Path, required=True)
    parser.add_argument("--candidate", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--file-only", action="store_true")
    args = parser.parse_args()
    before, after = summarize(args.baseline, args.file_only), summarize(args.candidate, args.file_only)
    if before.keys() != after.keys() or not before:
        raise ValueError("The two codec manifests differ or are empty")
    rows, files = [], []
    for key in sorted(before):
        a, b = before[key], after[key]
        rows.append(dict(Elements=key[0], States=key[1], Action=key[2], Before=a, After=b,
                         ReductionPercent=100 * (a["Median"] - b["Median"]) / a["Median"]))
    for elements, states in sorted({key[:2] for key in before}):
        name = f"Payload_{elements}_{states}.ndm"
        a, b = (args.baseline / name).read_bytes(), (args.candidate / name).read_bytes()
        if a != b:
            raise ValueError(f"The NDMSNP04 file changed: {name}")
        files.append(dict(Elements=elements, States=states, Bytes=len(a),
                          SHA256=hashlib.sha256(a).hexdigest().upper(), ByteIdentical=True))
    result = dict(Status="PASS", SyntheticIOOnly=True, SolverCalls=0, ExactAllCells=True,
                  FileFormat="NDMSNP04 unchanged", Payloads=files, Comparisons=rows)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(dict(Status="PASS", Comparisons=len(rows), ByteIdenticalFiles=len(files))))


if __name__ == "__main__":
    main()
