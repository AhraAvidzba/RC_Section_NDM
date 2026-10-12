"""Сигнатуры реального декларативного API; нет ручного списка методов."""
from __future__ import annotations

import argparse
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


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--inventory", action="store_true")
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--semantics-check", action="store_true")
    parser.add_argument("--sync-example", action="store_true")
    parser.add_argument("--render-draft", action="store_true")
    args = parser.parse_args()
    actual = inventory()
    if args.render_draft:
        from SectionRecipeApiText import purpose, parameter_meaning, LIFECYCLE, ERRORS
        lines = ["# SECTION_RECIPE_API: каталог текущих сигнатур", "",
                 "Стадия: DRAFT. Сигнатуры, defaults, параметры и жизненный цикл синхронизированы с source; минимальные компилируемые примеры/ссылки и native Help ещё проверяются. Это не закрытый E04/E12.", "",
                 actual["scope"], "", "Внутри мм/мм²/Н/Н*мм/МПа; углы в радианах. ByRef/Optional/default приведены в точной сигнатуре. Полный общий фасад и исходный класс представлены в ADDING_SECTION_TYPE.md.", ""]
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
            if member.get("resultType"):
                lines.extend(["Результат: " + member["resultType"] + ". Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.", ""])
            lines.extend(["Жизненный цикл: " + LIFECYCLE[owner], "",
                          "Ошибки/исправление: " + ERRORS.get(owner, "Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом."), "",
                          "Source: " + member["source"] + ":" + str(member["line"]), ""])
        (DOCS / "SECTION_RECIPE_API.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
        print(json.dumps({"status": "DOCUMENTATION_DRAFT", "members": len(actual["members"]), "missingSemantics": 0}))
        return
    if args.sync_example:
        guide = DOCS / "ADDING_SECTION_TYPE.md"
        example = DOCS / "examples/CSectionRecipeRoundedTrapezoid.cls"
        content = guide.read_text(encoding="utf-8")
        pattern = r"(?s)(<!-- complete-example:start -->).*?(<!-- complete-example:end -->)"
        if len(re.findall(pattern, content)) != 1: raise ValueError("Complete-example block is not unique")
        code = example.read_text(encoding="utf-8-sig").rstrip()
        guide.write_text(re.sub(pattern, lambda match: match[1] + "\n```vb\n" + code + "\n```\n" + match[2], content), encoding="utf-8")
        print(json.dumps({"status": "COMPLETE_CLASS_SYNCHRONIZED", "sourceSHA256": hashlib.sha256(example.read_bytes()).hexdigest()}))
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
        for key in ("members", "sourceHashes"):
            if saved[key] != actual[key]: raise ValueError(f"API manifest is stale: {key}")
        print(json.dumps({"status": "SOURCE_MANIFEST_MATCH", "members": len(actual["members"])}))
    else:
        target.write_text(json.dumps(actual, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(json.dumps({"status": actual["status"], "members": len(actual["members"])}))


if __name__ == "__main__": main()
