"""Сохраняет неизменяемый пакет геометрического этапа и проверяет каждый файл."""
from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--evidence", required=True)
    parser.add_argument("--archive", required=True)
    parser.add_argument("--extra-evidence", action="append", default=[])
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[2]
    archive = (root / args.archive).resolve()
    if not archive.is_relative_to(root / "docs/regression"):
        raise ValueError("Evidence archive must stay inside docs/regression")
    if archive.exists():
        raise FileExistsError("Existing evidence is immutable: " + str(archive))
    files: set[Path] = set()
    for folder in [args.evidence, *args.extra_evidence]:
        files.update(p for p in (root / folder).rglob("*") if p.is_file() and
                     p.suffix.lower() in {".txt", ".json", ".xlsm", ".png"})
    for folder in ["src", "tests", "docs/geometry_refactor/examples"]:
        files.update(p for p in (root / folder).rglob("*") if p.is_file() and
                     p.suffix.lower() in {".bas", ".cls", ".frm", ".frx", ".py"})
    files.update(p for p in (root / "docs/geometry_refactor").glob("*")
                 if p.is_file() and p.suffix.lower() in {".md", ".json"})
    for name in ["AGENTS.md", "README.md", "tools/build_workbook/Update-PrivateWorkbookCode.ps1",
                 "tools/build_workbook/Verify-GeometryBaselineCode.ps1", "tools/build_workbook/Run-Audit03Watchdog.ps1",
                 "tools/build_workbook/Run-Audit02ModeTests.ps1", "tools/build_workbook/Compare-GeometrySnapshots.py",
                 "tools/build_workbook/Inventory-GeometryEvidence.py", "tools/build_workbook/Pack-GeometryMilestone.py",
                 "tools/build_workbook/GeometryExampleSupport.ps1",
                 "tools/build_workbook/Build-SectionRecipeApi.py", "tools/build_workbook/SectionRecipeApiText.py",
                 "tools/build_workbook/Build_SectionRecipeApiSupport.py", "tools/build_workbook/Build-SectionRecipeApiExamples.py",
                 "tools/build_workbook/Build-SectionRecipeGuide.py", "tools/build_workbook/SectionRecipeGuide.ps1",
                 "tools/build_workbook/Update-SectionRecipeGuide.ps1", "tools/build_workbook/Verify-SectionRecipeGuide.py",
                 "tools/build_workbook/Summarize-SectionRecipeGuide.py",
                 "tools/build_workbook/Build-Workbook.ps1"]:
        files.add(root / name)
    entries = []
    archive.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(archive, "x", ZIP_DEFLATED) as out:
        for path in sorted(files):
            data = path.read_bytes()
            name = path.relative_to(root).as_posix()
            out.writestr(name, data)
            entries.append({"path": name, "bytes": len(data), "sha256": digest(data)})
    # Сверка распакованных байтов, а не только CRC или наличия имён.
    with ZipFile(archive) as check:
        for entry in entries:
            assert digest(check.read(entry["path"])) == entry["sha256"], entry["path"]
    manifest = {"schema": 1, "status": "ARCHIVE_INTEGRITY_PASS",
                "parentCommit": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=root, text=True).strip(),
                "archive": archive.relative_to(root).as_posix(), "archiveSHA256": digest(archive.read_bytes()),
                "entries": entries}
    archive.with_suffix(".manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2)+"\n", encoding="utf-8")
    print(json.dumps({"status": manifest["status"], "entries": len(entries),
                      "archiveSHA256": manifest["archiveSHA256"]}))


if __name__ == "__main__":
    main()
