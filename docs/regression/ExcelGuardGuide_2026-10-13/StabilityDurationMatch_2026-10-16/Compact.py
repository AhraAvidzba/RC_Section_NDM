"""Архивирует доказательства и данные листов, затем убирает свои временные книги."""
import hashlib
import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

D = Path(__file__).resolve().parent


def sha(data):
    return hashlib.sha256(data).hexdigest().upper()


def receipt(name):
    return json.loads((D / name).read_text(encoding="utf-8-sig"))


assert receipt("Acceptance.json")["PublishedVerified"]
archive_path = D / "Evidence.zip"
assert not archive_path.exists()
names = ["BaselineNativeVBA.txt", "CandidateNativeVBA.txt", "UserNativeVBA.txt", "AcceptedNativeVBA.txt",
         "PublishedNativeVBA.txt", "ClassComments.json",
         "BeforeFix/modTestProfileConfig.RunStabilityDurationMatchingTests.txt",
         "Matching/modTestProfileConfig.RunStabilityDurationMatchingTests.txt",
         "DurationRegression/modTestProfileConfig.RunAudit03DurationConfigTests.txt"]
contents = {name: (D / name).read_bytes() for name in names}
# Для повторной проверки 106 тысяч ячеек достаточно исходных ZIP-частей без VBA.
for name in ("UserBaseline", "Accepted"):
    data = BytesIO()
    with ZipFile(D / (name + ".xlsm")) as original, ZipFile(data, "w") as bundle:
        for info in original.infolist():
            if info.filename != "xl/vbaProject.bin":
                bundle.writestr(info, original.read(info.filename))
    contents[name + ".data.zip"] = data.getvalue()
records = {name: sha(data) for name, data in contents.items()}
with ZipFile(archive_path, "w", ZIP_DEFLATED, compresslevel=9) as archive:
    for name, data in contents.items():
        archive.writestr(name, data)
with ZipFile(archive_path) as archive:
    assert archive.testzip() is None and set(archive.namelist()) == set(contents)
    assert all(sha(archive.read(name)) == expected for name, expected in records.items())
copies = {"Baseline.xlsm": receipt("Before.json")["Hashes"]["workbook/output/RC_Section_NDM.xlsm"],
          "UserBaseline.xlsm": receipt("UserSnapshot.json")["Hashes"]["workbook/output/RC_Section_NDM.xlsm"],
          "Accepted.xlsm": receipt("Package.json")["AcceptedSHA256"],
          "Candidate/Matching.xlsm": receipt("Matching/Run.json")["SourceSHA256"]}
for stage in ("BeforeFix", "Matching", "DurationRegression"):
    copies[stage + "/GeometrySolidProbe.xlsm"] = receipt(stage + "/Run.json")["SourceSHA256"]
removed = []
for name, expected in copies.items():
    path = (D / name).resolve()
    path.relative_to(D)
    assert path.suffix == ".xlsm" and sha(path.read_bytes()) == expected
    removed.append(dict(Path=name, Bytes=path.stat().st_size, SHA256=expected))
# Все конечные пути и контрольные суммы проверены до первого удаления.
for item in removed:
    (D / item["Path"]).unlink()
for name in names:
    (D / name).unlink()
result = dict(ArchiveSHA256=sha(archive_path.read_bytes()), ArchiveBytes=archive_path.stat().st_size,
              Entries=records, RemovedCopies=removed, RemovedWorkbookBytes=sum(x["Bytes"] for x in removed))
(D / "Archive.json").write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
print(json.dumps(result, indent=2))
