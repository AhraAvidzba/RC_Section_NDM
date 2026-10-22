"""Сверяет baseline/итог геометрии, битовые результаты НДМ и настоящие схемы Excel."""
from pathlib import Path
import hashlib
import json
from PIL import Image, ImageChops

ROOT = Path(__file__).resolve().parents[2]
S = ROOT / "docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14"


def digest(data):
    return hashlib.sha256(data).hexdigest()


def verify():
    errors, rows, bits, images = [], [], [], []
    for path in sorted((S / "BaselineGeometry").glob("*.txt")):
        after = S / "SolidGeometry" / path.name
        # ANCHOR — удалённый временный буфер layout. Готовые ANNOTATION не исключаются.
        source = path.read_bytes()
        baseline = b"\n".join(line for line in source.splitlines() if b"\tANCHOR\t" not in line)
        actual = b"\n".join(after.read_bytes().splitlines())
        equal = baseline == actual
        rows.append({"File": path.name, "Passed": equal, "RemovedScratchAnchors": sum(b"\tANCHOR\t" in line for line in source.splitlines()),
                     "ComparedBytes": len(baseline), "BeforeSHA256": digest(baseline), "AfterSHA256": digest(actual)})
        if not equal: errors.append("Geometry differs: " + path.name)
    if len(rows) != 6: errors.append("Expected six complete geometry snapshots")
    for name in ("TestDumpSolidPhysicalBits", "TestDumpSolidDirectNdsBits"):
        file = "modTestGeometrySolidSnapshots." + name + ".txt"
        before = (S / "BaselineBits" / file).read_bytes()
        after = (S / "SolidBits" / file).read_bytes()
        equal = before == after
        bits.append({"File": file, "Passed": equal, "Bytes": len(before), "BeforeSHA256": digest(before), "AfterSHA256": digest(after)})
        if not equal: errors.append("Bitwise values differ: " + name)
        if name == "TestDumpSolidDirectNdsBits":
            lines = after.decode("utf-8-sig").splitlines()
            if len(lines) != 120 or any(line.split("|")[3] != "True" for line in lines):
                errors.append("Expected 120 converged NDM solutions")
    file = "modTestShapeArchitecture.TestDumpShapePreviewSnapshot.txt"
    a, b = (S / "BaselinePreview" / file).read_bytes(), (S / "SolidPreview" / file).read_bytes()
    if a != b: errors.append("Excel chart/shape coordinates or labels differ")
    for path in sorted((S / "BaselinePreview").glob("*.png")):
        after = S / "SolidPreview" / path.name
        with Image.open(path) as first, Image.open(after) as second:
            a, b = first.convert("RGBA"), second.convert("RGBA")
            equal = a.size == b.size and ImageChops.difference(a, b).getbbox() is None
            # getbbox on RGBA alone can ignore RGB differences under equal alpha.
            equal = equal and a.tobytes() == b.tobytes()
            images.append({"File": path.name, "Passed": equal, "Size": list(a.size), "BeforePixelsSHA256": digest(a.tobytes()), "AfterPixelsSHA256": digest(b.tobytes())})
            if not equal: errors.append("Excel rendered image differs: " + path.name)
    if len(images) != 9: errors.append("Expected nine Excel PNG previews")
    result = {"Passed": not errors, "Geometry": rows, "BitwiseModelsAndNds": bits, "NdsConverged": 120,
              "ChartStructureEqual": (S / "BaselinePreview" / file).read_bytes() == (S / "SolidPreview" / file).read_bytes(),
              "ExcelPng": images, "Errors": errors}
    (S / "RegressionEquality.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"Passed": result["Passed"], "GeometrySnapshots": len(rows), "NdsSolutions": 120,
                      "ExcelPng": len(images), "Errors": errors}, ensure_ascii=False))
    if errors: raise SystemExit(1)


if __name__ == "__main__": verify()
