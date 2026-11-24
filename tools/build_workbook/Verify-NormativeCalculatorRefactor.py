"""ДЛЯ ТЕСТОВ: проверяет сохранность выражений при выделении нормативных классов.

Не заменяет нативные расчёты: сравнивает тела методов после удаления только
квалификаторов общего слоя и явного аргумента состояния, сохраняя строки и числа.
"""
import argparse
import json
from pathlib import Path
import re
import subprocess


def methods(text):
    """Читает методы VBA, отделяя продолжение сигнатуры от исполняемого тела."""
    result = {}
    for match in re.finditer(r"(?m)^(?:Public|Private|Friend) (Sub|Function|Property (?:Get|Let|Set)) (\w+)\(", text):
        kind, name = match.groups()
        end_kind = "Property" if kind.startswith("Property") else kind
        end = re.search(r"(?m)^End " + end_kind + r"\b", text[match.end():])
        signature_end = text.index("\n", match.end())
        while text[:signature_end].rstrip().endswith("_"):
            signature_end = text.index("\n", signature_end + 1)
        result[(name, kind)] = text[signature_end:match.end() + end.start()]
    return result


CRACK_COMMON = "StrainAtConcreteCenter UtilizationFromData SetWidthResult UserStopReason SetWidthInvalidInput FinishNoCrack AppendDiagnostic FormatNumberInvariant MinDouble MaxDouble".split()
STABILITY_COMMON = "ConcreteFaceProjectionForMoment MomentAboutProjection EnsurePositivePlaneGeometry StabilityNumber SignForEccentricity InitialEccentricity PhiLByMode PhiLFromMomentRatio EtaByCriticalForce ReserveByCriticalForce DesignMomentWithEta".split()


def canonical(text):
    """Нормализует только перенос методов, без округления чисел и изменения строк."""
    tokens = []
    aliases = {"crack" + n.lower(): n.lower() for n in CRACK_COMMON}
    aliases.update({"stability" + n.lower(): n.lower() for n in STABILITY_COMMON})
    aliases["stabilityraiseinput"] = "raiseinput"
    aliases["tstabilityplane"] = "tplaneresult"
    for part in re.findall(r'"(?:[^"\n]|"")*"|\x27[^\n]*|[^"\x27]+', text):
        if part.startswith("'"):
            continue
        if part.startswith('"'):
            tokens.append(part)
            continue
        part = re.sub(r"\b(?:mWork|work)\.(\w+)", lambda m: "mAs" if m[1] == "RebarArea" else "m" + m[1], part)
        part = re.sub(r"\b(?:mSP35|mSP63)\.", "", part)
        part = re.sub(r"\b(?:mWork|work)\s*,\s*", "", part)
        part = re.sub(r"\b(?:normBranch|normTableInputError)\s*,\s*", "", part)
        part = re.sub(r"\bnormBranch\b", "mBranch", part)
        part = re.sub(r"\bnormTableInputError\b", "mSP35TableInputError", part)
        part = re.sub(r"\b(?:mWork|work)(?=\s*\))", "", part)
        part = re.sub(r"\b(?:CRACK_|STABILITY_)(\w+)", r"\1", part)
        part = re.sub(r"\b\w+\b", lambda m: aliases.get(m[0].lower(), m[0].lower()), part)
        tokens.extend(re.findall(r"[A-Za-z_]\w*|\d+(?:\.\d*)?(?:[Ee][+-]?\d+)?|[^\s_]", part))
    return tokens


def verify(root, baseline):
    """Сверяет каждый перенесённый метод с его единственной исходной реализацией."""
    checks = []
    for family, facade, shared in [
        ("Crack", "CCrackWidthCalculator", "modCrackWidthSupport"),
        ("Stability", "CStabilityCalculator", "modStabilityPlane"),
    ]:
        path = f"src/{family}/{facade}.cls"
        old_text = subprocess.check_output(["git", "show", baseline + ":" + path], cwd=root).decode("utf-8-sig")
        old = methods(old_text)
        prefix = "CCrackWidth" if family == "Crack" else "CStability"
        paths = [root / f"src/{family}/{prefix}{code}Calculator.cls" for code in ["SP35", "SP63"]]
        paths.append(root / f"src/{family}/{shared}.bas")
        shared_text = paths[-1].read_text(encoding="utf-8-sig")
        old_constants = re.findall(r"(?m)^Private Const ([^\n]+)", old_text)
        new_constants = re.findall(r"(?m)^Public Const (?:CRACK_|STABILITY_)([^\n]+)", shared_text)
        assert [canonical(v) for v in old_constants] == [canonical(v) for v in new_constants], f"Changed {family} constants"
        if family == "Stability":
            old_type = re.search(r"(?ms)^Private Type TPlaneResult\n(.*?)^End Type", old_text)[1]
            new_type = re.search(r"(?ms)^Public Type TStabilityPlane\n(.*?)^End Type", shared_text)[1]
            assert canonical(old_type) == canonical(new_type), "Changed stability plane fields"
        else:
            original_fields = []
            for line in re.findall(r"(?m)^Private m\w+ As [^\n]+", old_text):
                for entry in line.split("'")[0].removeprefix("Private ").split(", "):
                    original_fields.append(canonical(entry))
            new_type = re.search(r"(?ms)^Public Type TCrackWidthWork\n(.*?)^End Type", shared_text)[1]
            current_fields = []
            for entry in new_type.splitlines():
                entry = entry.split("'")[0].strip()
                if not entry:
                    continue
                name, rest = entry.split(" ", 1)
                current_fields.append(canonical(("mAs" if name == "RebarArea" else "m" + name) + " " + rest))
            assert original_fields == current_fields, "Changed crack state fields"
        transferred = set()
        for target in paths:
            text = target.read_text(encoding="utf-8-sig")
            for (name, kind), body in methods(text).items():
                original_name = name
                for n in CRACK_COMMON + STABILITY_COMMON:
                    if name in ("Crack" + n, "Stability" + n):
                        original_name = n
                if name in ("TableInputError", "Branch", "StabilityRaiseInput"):
                    continue
                key = (original_name, kind)
                if key not in old:
                    raise AssertionError(f"No original method: {target.name}.{name}")
                expected = old[key]
                if name == "CalculateSP63":
                    expected = re.sub(r"(?m)^\s*mBranch = CODE_SP63\s*$", "", expected)
                actual_tokens, expected_tokens = canonical(body), canonical(expected)
                if actual_tokens != expected_tokens:
                    import difflib
                    difference = list(difflib.unified_diff(expected_tokens, actual_tokens))[:50]
                    raise AssertionError(f"Changed calculation: {target.name}.{name}\n" + "\n".join(difference))
                checks.append({"Component": target.stem, "Method": name, "ExpressionsAndLiteralsUnchanged": True})
                transferred.add(key)
        current_facade = methods((root / path).read_text(encoding="utf-8-sig"))
        for key, body in old.items():
            if key in transferred:
                continue
            if key not in current_facade:
                raise AssertionError(f"Lost facade method: {facade}.{key}")
            actual = current_facade[key]
            if key[0] == "Class_Initialize" and family == "Crack":
                actual = re.sub(r"(?m)^\s*Set mSP(?:35|63) = New CCrackWidthSP(?:35|63)Calculator\s*$", "", actual)
            if canonical(actual) != canonical(body):
                raise AssertionError(f"Changed facade preparation: {facade}.{key}")
            checks.append({"Component": facade, "Method": key[0], "ExpressionsAndLiteralsUnchanged": True})
    return checks


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--baseline-ref", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    project = Path(__file__).resolve().parents[2]
    checks = verify(project, args.baseline_ref)
    test_path = "tests/modTestBatchCalculation.bas"
    original_tests = methods(subprocess.check_output(["git", "show", args.baseline_ref + ":" + test_path], cwd=project).decode("utf-8-sig"))
    current_tests = methods((project / test_path).read_text(encoding="utf-8-sig"))
    assert current_tests.keys() - original_tests.keys() == {("RunNormativeStabilityTests", "Function"), ("RunNormativeGuardedTest", "Function"), ("RunNormativePartialBranchFailureTest", "Function")}
    for key, body in original_tests.items():
        assert key in current_tests and canonical(body) == canonical(current_tests[key]), f"Changed original test: {key}"
    output = {"Status": "PASS", "BaselineRef": args.baseline_ref, "Methods": len(checks), "Checks": checks,
              "ExistingBatchTestMethodsUnchanged": len(original_tests), "NewTestEntrypointsOnly": True,
              "Limit": "Expression preservation; native calculations and diagnostics require directed acceptance."}
    (project / args.output).write_text(json.dumps(output, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({k: v for k, v in output.items() if k != "Checks"}, ensure_ascii=False))
