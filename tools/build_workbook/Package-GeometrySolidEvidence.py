"""Компактный архив полных снимков и native VBA без временных XLSM и пользовательских данных."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import hashlib
import json

ROOT = Path(__file__).resolve().parents[2]
S = ROOT / "docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14"
paths = [S / name for name in ("BaselineNativeVBA.txt", "TestedNativeVBA.txt", "FinalNativeVBA.txt")]
for stage in ("BaselineBits", "SolidBits", "SolidGeometry", "BaselinePreview", "SolidPreview"):
    for path in sorted((S / stage).iterdir()):
        if path.suffix in (".txt", ".png") or path.name == "Run.json": paths.append(path)
with ZipFile(S / "RegressionSnapshots.zip", "w", ZIP_DEFLATED, compresslevel=9) as archive:
    for path in paths: archive.write(path, path.relative_to(S).as_posix())
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
receipt = {"Archive": "RegressionSnapshots.zip", "ArchiveSHA256": sha(S / "RegressionSnapshots.zip"),
           "Files": {path.relative_to(S).as_posix(): sha(path) for path in paths}}
(S / "EvidencePackage.json").write_text(json.dumps(receipt, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
print(json.dumps({"PackedFiles": len(paths), "Bytes": (S / "RegressionSnapshots.zip").stat().st_size}))
