"""Содержательные инструкции/ветвления: не считает комментарии, объявления и переносы строк."""
from pathlib import Path
import hashlib
import json
import re
import subprocess

ROOT = Path(__file__).resolve().parents[2]
S = ROOT / "docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14"
RECIPES = ["src/Geometry/CSectionRecipe" + name + ".cls" for name in ("Circle", "Rectangles", "PointedSection", "HollowRectangle")]
RECIPES += ["docs/geometry_refactor/examples/CSectionRecipeRoundedTrapezoid.cls"]
COMMON = ["src/Geometry/CPolygonSpec.cls", "src/Section/CSectionDefinitionBuilder.cls", "src/Section/CSectionDefinition.cls",
          "src/Section/CSectionGeometryQuery.cls", "src/Section/CRebarRuleOptions.cls", "src/Section/CSectionInputTable.cls", "src/Common/modSectionRecipeTypes.bas"]


def measure(text):
    instructions, branches, methods, active, pending = 0, 0, 0, False, ""
    for physical in text.splitlines():
        tokens = re.split(r'("(?:[^"\n]|"")*"|\x27[^\n]*)', physical)
        code = "".join(token for token in tokens if not token.startswith("'"))
        line = pending + code.strip()
        if line.endswith(" _"):
            pending = line[:-2] + " "
            continue
        pending = ""
        if re.match(r"(?:Public|Private|Friend) (?:Function|Sub|Property Get)\b", line, re.I):
            active = True
            methods += 1
            # One-line getters still have an executable assignment after the colon.
            line = re.sub(r"^[^:]*", "", line)
        if not active: continue
        chunks = re.split(r'("(?:[^"\n]|"")*")', line)
        masked = "".join('""' if i % 2 else token for i, token in enumerate(chunks))
        for statement in masked.split(":"):
            statement = statement.strip()
            if not statement: continue
            if re.match(r"End (?:Function|Sub|Property)\b", statement, re.I): active = False; continue
            if re.match(r"(?:Dim|Static|Const|End If|End Select|Next\b|Loop\b|Wend\b|End With|Else$|Case\b)\b", statement, re.I): continue
            if re.fullmatch(r"\w+", statement) and statement in ("Finished", "Failed", "InvalidGeometry", "NextFace", "NextSide", "InvalidPath"): continue
            instructions += 1
            if re.match(r"(?:If\b|ElseIf\b|Select Case\b)", statement, re.I):
                branches += 1
                if re.search(r"\bThen\s+\S", statement, re.I): instructions += 1
    return {"Instructions": instructions, "Branches": branches, "Methods": methods}


if __name__ == "__main__":
    before_path = S / "SourceMetricsBefore.json"
    if not before_path.exists():
        head = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip()
        values = {}
        for name in RECIPES + COMMON:
            result = subprocess.run(["git", "show", head + ":" + name], cwd=ROOT, capture_output=True)
            text = result.stdout.decode("utf-8-sig") if result.returncode == 0 else ""
            values[name] = measure(text)
        before = {"Head": head, "Method": "Procedure-body statements; declarations/comments/closing delimiters excluded; inline Then counted separately; colon-separated statements counted individually", "Sources": values}
        before_path.write_text(json.dumps(before, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    before = json.loads(before_path.read_text(encoding="utf-8"))
    result = {"BaselineHead": before["Head"], "Method": before["Method"], "Recipes": {}, "Common": {}}
    for name in RECIPES + COMMON:
        current = (ROOT / name).read_text(encoding="utf-8-sig") if (ROOT / name).exists() else ""
        result["Recipes" if name in RECIPES else "Common"][name] = {"Before": before["Sources"][name], "After": measure(current), "SourceSHA256": hashlib.sha256(current.encode("utf-8")).hexdigest()}
    (S / "SourceMetrics.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result["Recipes"], ensure_ascii=False))
