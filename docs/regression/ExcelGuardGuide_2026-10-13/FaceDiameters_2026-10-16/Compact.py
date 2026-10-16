"""Сохраняет проверяемые журналы и удаляет только три собственные копии книги."""
import hashlib
import json
import zipfile
from pathlib import Path

directory = Path(__file__).resolve().parent


def digest(data):
    return hashlib.sha256(data).hexdigest().upper()


def read_json(relative):
    return json.loads((directory / relative).read_text(encoding="utf-8-sig"))


archive_path = directory / "Evidence.zip"
assert not archive_path.exists(), "Archive already exists"
logs = ["Native/modTestSectionRecipes.DumpFaceDiameterProbe.txt",
        "Directed/modTestSectionRecipes.DumpFaceDiameterProbe.txt"]
records = {name: digest((directory / name).read_bytes()) for name in logs}
with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
    for name in logs:
        archive.write(directory / name, name)
with zipfile.ZipFile(archive_path) as archive:
    assert archive.testzip() is None
    assert set(archive.namelist()) == set(logs)
    for name, expected in records.items():
        assert digest(archive.read(name)) == expected
prep = read_json("Preparation.json")
copies = {"Candidate/FaceDiameters.xlsm": prep["CandidateSHA256"],
          "Native/GeometrySolidProbe.xlsm": read_json("Native/Run.json")["SourceSHA256"],
          "Directed/GeometrySolidProbe.xlsm": read_json("Directed/Run.json")["SourceSHA256"]}
removed = []
for name, expected in copies.items():
    path = (directory / name).resolve()
    path.relative_to(directory)
    assert path.suffix == ".xlsm" and digest(path.read_bytes()) == expected
    removed.append(dict(path=name, bytes=path.stat().st_size, SHA256=expected))
# Проверка всех путей и контрольных сумм выполнена до первого удаления.
for item in removed:
    (directory / item["path"]).unlink()
for name in logs:
    (directory / name).unlink()
manifest = dict(ArchiveSHA256=digest(archive_path.read_bytes()), ArchiveBytes=archive_path.stat().st_size,
                Logs=records, RemovedPrivateCopies=removed, RemovedWorkbookBytes=sum(x["bytes"] for x in removed))
(directory / "Archive.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
print(json.dumps(manifest, indent=2))
