"""Сигнатуры реального декларативного API; нет ручного списка методов."""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DOCS = ROOT / "docs/geometry_refactor"
SOURCES = [
    "src/Section/ISectionRecipe.cls", "src/Section/CSectionDefinitionBuilder.cls",
    "src/Geometry/CGeometryPath.cls", "src/Section/CSectionDefinition.cls",
    "src/Section/CRebarRuleOptions.cls", "src/Geometry/CGeometryDiagnostic.cls",
    "src/Section/CSectionInputTable.cls", "src/Section/CSectionRecipeContext.cls",
    "src/Excel/CSectionInputSnapshotReader.cls", "src/Excel/CSectionTypeRegistry.cls",
    "src/Section/CSectionBuilder.cls", "src/Section/CRebarRuleExecutor.cls",
    "src/Excel/CUnitSystem.cls", "src/Section/CConcreteRegion.cls",
    "src/Section/CSectionGeometryQuery.cls", "src/Common/modSectionRecipeTypes.bas",
    "src/Common/modGeometryTypes.bas",
]


def split_top(text: str, separator=",") -> list[str]:
    result, start, depth, quoted, i = [], 0, 0, False, 0
    while i < len(text):
        character = text[i]
        if character == '"':
            if quoted and i + 1 < len(text) and text[i + 1] == '"':
                i += 2
                continue
            quoted = not quoted
        elif not quoted:
            if character == "(": depth += 1
            elif character == ")": depth -= 1
            elif character == separator and depth == 0:
                result.append(text[start:i].strip()); start = i + 1
        i += 1
    result.append(text[start:].strip())
    return result


def declarations(path: Path) -> list[dict]:
    source = path.read_text(encoding="utf-8-sig")
    owner = re.search(r'Attribute VB_Name = "([^"]+)"', source).group(1)
    logical, pending, first = [], "", 0
    for number, line in enumerate(source.splitlines(), 1):
        if not pending: first = number
        pending += line.strip()
        if pending.endswith(" _"):
            pending = pending[:-2] + " "
        else:
            logical.append((first, pending)); pending = ""
    members, block = [], None
    for number, line in logical:
        enum = re.match(r"Public (Enum|Type) (\w+)$", line)
        if enum:
            block = {"owner": owner, "name": enum[2], "kind": enum[1], "signature": line,
                     "fields": [], "line": number, "source": path.relative_to(ROOT).as_posix()}
            members.append(block)
            continue
        if block is not None:
            if line == "End " + block["kind"]: block = None
            elif line and not line.startswith("'"):
                block["fields"].append(line.split("'", 1)[0].strip())
            continue
        if not line.startswith("Public "): continue
        line = split_top(line, ":")[0]
        match = re.match(r"Public (Function|Sub|Property Get|Property Let|Property Set) (\w+)\((.*)\)(?: As ([\w]+))?$", line)
        if match:
            arguments = []
            for argument in split_top(match[3]):
                if not argument: continue
                param = re.match(r"(?:(Optional) )?(?:(ByVal|ByRef|ParamArray) )?(\w+)(\(\))? As (\w+)(?: = (.*))?$", argument)
                if not param: raise ValueError(f"Unparsed parameter: {owner}.{match[2]}: {argument}")
                arguments.append({"name": param[3], "type": param[5] + (param[4] or ""),
                                  "passing": param[2] or "ByRef", "optional": bool(param[1]), "default": param[6], "signature": argument})
            members.append({"owner": owner, "name": match[2], "kind": match[1], "signature": line,
                            "parameters": arguments, "resultType": match[4], "line": number,
                            "source": path.relative_to(ROOT).as_posix()})
        else:
            field = re.match(r"Public (\w+) As (\w+)$", line)
            if field:
                members.append({"owner": owner, "name": field[1], "kind": "Field", "signature": line,
                                "resultType": field[2], "line": number, "source": path.relative_to(ROOT).as_posix()})
            else:
                constant = re.match(r"Public Const (\w+) As (\w+) = (.*)", line)
                if constant:
                    members.append({"owner": owner, "name": constant[1], "kind": "Const", "signature": line,
                                    "resultType": constant[2], "value": constant[3], "line": number,
                                    "source": path.relative_to(ROOT).as_posix()})
                else: raise ValueError(f"Unparsed public member: {owner}: {line}")
    for member in members:
        member["id"] = member["owner"] + "." + member["name"]
        member["anchor"] = "api-" + member["id"].lower().replace(".", "-")
    return members


def inventory() -> dict:
    members = []
    for source in SOURCES: members.extend(declarations(ROOT / source))
    ids = [member["id"] for member in members]
    if len(ids) != len(set(ids)): raise ValueError("Duplicate API ID")
    return {"schema": 1, "status": "SOURCE_INVENTORY_ONLY", "members": members,
            "scope": "Все Public члены фасада, его options/результатов и общего ввода/единиц; общие query/region и build diagnostics для проверок. Private/Friend не API рецепта.",
            "sourceHashes": {source: hashlib.sha256((ROOT / source).read_bytes()).hexdigest() for source in SOURCES}}


def verify_catalog(actual: dict, saved: dict, content: str, examples: dict):
    for key in ("members", "sourceHashes"):
        if saved[key] != actual[key]: raise ValueError("API manifest is stale: " + key)
    anchors = re.findall(r'<a id="([^"]+)"></a>', content)
    if anchors != [member["anchor"] for member in actual["members"]]: raise ValueError("Catalog members/anchors differ")
    for member in actual["members"]:
        pattern = r'(?s)<a id="' + re.escape(member["anchor"]) + r'"></a>\s*### ' + re.escape(member["id"]) + r'\s+(.*?)(?=<a id=|\Z)'
        block = re.search(pattern, content)
        if not block or "```vb\n" + member["signature"] + "\n```" not in block[1]:
            raise ValueError("Catalog signature differs: " + member["id"])
        example = examples["examples"].get(member["id"])
        if not example or not example["nativeRoots"]: raise ValueError("Missing verified example: " + member["id"])
        lines = (ROOT / example["source"]).read_text(encoding="utf-8-sig").splitlines()
        raw = "\n".join(lines[example["line"]-1:example["line"]-1+len(example["code"].splitlines())])
        if raw != example["code"]: raise ValueError("Example source moved/changed: " + member["id"])
        if "```vb\n" + example["code"] + "\n```" not in block[1]: raise ValueError("Example text differs: " + member["id"])
    for collection in ("sourceHashes", "reportHashes"):
        for name, expected in examples[collection].items():
            if hashlib.sha256((ROOT/name).read_bytes()).hexdigest() != expected:
                raise ValueError("Example/evidence changed: " + name)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--inventory", action="store_true")
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--semantics-check", action="store_true")
    parser.add_argument("--sync-example", action="store_true")
    parser.add_argument("--render-draft", action="store_true")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    actual = inventory()
    if args.self_test:
        saved = json.loads((DOCS / "SECTION_RECIPE_API_MANIFEST.json").read_text(encoding="utf-8-sig"))
        content = (DOCS / "SECTION_RECIPE_API.md").read_text(encoding="utf-8-sig")
        examples = json.loads((DOCS / "SECTION_RECIPE_API_EXAMPLES.json").read_text(encoding="utf-8-sig"))
        verify_catalog(actual, saved, content, examples)
        altered = copy.deepcopy(saved); altered["members"][0]["signature"] += " As StaleSignature"
        probes = [(altered, content), (saved, content.replace(actual["members"][0]["signature"], "Public Sub RemovedMethod()", 1)),
                  (saved, content.replace('<a id="'+actual["members"][0]["anchor"]+'"></a>', '', 1))]
        detected = 0
        for candidate, text in probes:
            try: verify_catalog(actual, candidate, text, examples)
            except ValueError: detected += 1
        if detected != len(probes): raise ValueError("A stale manifest/doc mutation was accepted")
        print(json.dumps({"status": "API_CATALOG_MUTATION_GUARDS_PASS", "members": len(actual["members"]), "detected": detected}))
        return
    if args.render_draft:
        from SectionRecipeApiText import purpose, parameter_meaning, LIFECYCLE, ERRORS
        examples_path = DOCS / "SECTION_RECIPE_API_EXAMPLES.json"
        examples = json.loads(examples_path.read_text(encoding="utf-8-sig")) if examples_path.exists() else {"examples": {}}
        option_source = (ROOT / "src/Section/CRebarRuleOptions.cls").read_text(encoding="utf-8-sig")
        initialization = re.search(r"(?s)Private Sub Class_Initialize\(\)(.*?)End Sub", option_source).group(1)
        explicit_defaults = dict(re.findall(r"(?m)^\s*(\w+) = ([^\r\n]+)", initialization))
        ready = examples.get("status") == "SOURCE_CALLS_AND_NATIVE_ROOTS_MATCH"
        status = "Сигнатуры, defaults и каждый контекстный вызов синхронизированы с source; native entrypoint/журнал указан у каждого члена. Workbook-расположение фиксирует NativeMap.json, итоговая приёмка книги — отдельный отчёт." if ready else "Стадия: DRAFT. Примеры/ссылки ещё сопоставляются; E04/E12 остаются открытыми."
        lines = ["# SECTION_RECIPE_API: каталог текущих сигнатур", "", status, "",
                 actual["scope"], "", "Внутри мм/мм²/Н/Н*мм/МПа; углы в радианах. ByRef/Optional/default приведены в точной сигнатуре. Краткий порядок подключения — в ADDING_SECTION_TYPE.md; полный API — в modSectionRecipeApiGuide, тела методов — в исходных классах VBE.", ""]
        owner = None
        for member in actual["members"]:
            if member["owner"] != owner:
                owner = member["owner"]
                lines.extend(["## " + owner, "", LIFECYCLE[owner], ""])
            lines.extend([f'<a id="{member["anchor"]}"></a>', "### " + member["id"], "", purpose(member), "",
                          "```vb", member["signature"], "```", ""])
            if member.get("parameters"):
                lines.extend(["| Параметр | Тип / передача / default | Смысл и единицы |", "|---|---|---|"])
                for parameter in member["parameters"]:
                    flags = parameter["type"] + "; " + parameter["passing"]
                    flags += "; Optional" if parameter["optional"] else "; обязательный"
                    if parameter["default"] is not None: flags += "; default=" + parameter["default"]
                    elif parameter["optional"]: flags += "; пропущенный Variant=Missing"
                    lines.append("| " + parameter["name"] + " | " + flags.replace("|", "\\|") + " | " + parameter_meaning(member, parameter).replace("|", "\\|") + " |")
                lines.append("")
            if member.get("fields"):
                lines.extend(["Поля/селекторы по source:", "", "```vb", *member["fields"], "```", ""])
            if member["owner"] == "CRebarRuleOptions":
                implicit = {"String": "vbNullString", "Double": "0#", "Boolean": "False", "Long": "0"}
                default = explicit_defaults.get(member["name"], implicit.get(member["resultType"], "0"))
                lines.extend(["Default после New: " + default + ". Поле изменяется до объявления правила; facade копирует его значение.", ""])
            if member.get("resultType"):
                label = "Тип записываемого поля" if member["kind"] == "Field" else "Тип результата"
                lines.extend([label + ": " + member["resultType"] + ". Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.", ""])
            lines.extend(["Жизненный цикл: " + LIFECYCLE[owner], "",
                          "Ошибки/исправление: " + ERRORS.get(owner, "Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом."), "",
                          "Source: " + member["source"] + ":" + str(member["line"]), ""])
            example = examples["examples"].get(member["id"])
            if example:
                lines.extend(["Минимальный вызов; требуемый контекст: " + example["context"], "",
                              "```vb", example["code"].strip("\n"), "```", "",
                              "Полный контекст: " + example["source"] + ":" + str(example["line"]) + ".", ""])
                for proof in example["nativeRoots"]:
                    lines.extend(["Native entrypoint: " + proof["entryPoint"] + "; журнал: " + proof["reports"][-1] + ".", ""])
                if not example["nativeRoots"]:
                    lines.extend(["Статус примера: вызов найден; native coverage ещё не подтверждено.", ""])
            else:
                lines.extend(["Статус примера: проверенный вызов ещё не сопоставлен. Требование каталога остаётся открытым.", ""])
        (DOCS / "SECTION_RECIPE_API.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
        print(json.dumps({"status": "DOCUMENTATION_CALLS_SYNCHRONIZED" if ready else "DOCUMENTATION_DRAFT", "members": len(actual["members"]), "missingSemantics": 0}))
        return
    if args.sync_example:
        # Новая цель пользователя запрещает встраивать тела классов/методов.
        # Флаг сохранён для прежних команд, но полного дубликата больше не создаёт.
        print(json.dumps({"status": "EXAMPLES_REFERENCED_IN_SOURCE_ONLY", "embeddedBodies": 0}))
        return
    if args.semantics_check:
        from SectionRecipeApiText import purpose, parameter_meaning
        missing = []
        for member in actual["members"]:
            try: purpose(member)
            except ValueError as error: missing.append(str(error))
            for parameter in member.get("parameters", []):
                try: parameter_meaning(member, parameter)
                except ValueError as error: missing.append(str(error))
        print(json.dumps({"members": len(actual["members"]), "missing": sorted(set(missing))}, ensure_ascii=False, indent=2))
        return
    target = DOCS / "SECTION_RECIPE_API_MANIFEST.json"
    if args.check:
        saved = json.loads(target.read_text(encoding="utf-8-sig"))
        content = (DOCS / "SECTION_RECIPE_API.md").read_text(encoding="utf-8-sig")
        examples = json.loads((DOCS / "SECTION_RECIPE_API_EXAMPLES.json").read_text(encoding="utf-8-sig"))
        verify_catalog(actual, saved, content, examples)
        print(json.dumps({"status": "SOURCE_MANIFEST_DOC_EXAMPLES_MATCH", "members": len(actual["members"])}))
    else:
        target.write_text(json.dumps(actual, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(json.dumps({"status": actual["status"], "members": len(actual["members"])}))


if __name__ == "__main__": main()
