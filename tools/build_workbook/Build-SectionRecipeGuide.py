"""Один источник workbook-справки: строки Markdown и точные строки полного .cls."""
from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DOCS = ROOT / "docs/geometry_refactor"


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def build() -> dict:
    sources = ["ADDING_SECTION_TYPE.md", "SECTION_RECIPE_API.md", "SECTION_RECIPE_API_MANIFEST.json",
               "examples/CSectionRecipeRoundedTrapezoid.cls"]
    guide = (DOCS / sources[0]).read_text(encoding="utf-8-sig")
    example = (DOCS / sources[3]).read_text(encoding="utf-8-sig").rstrip()
    if "```vb\n" + example + "\n```" not in guide:
        raise ValueError("Complete .cls in guide is stale")
    manifest = json.loads((DOCS / sources[2]).read_text(encoding="utf-8-sig"))
    for source, expected in manifest["sourceHashes"].items():
        if sha(ROOT / source) != expected:
            raise ValueError("API source is stale: " + source)
    rows, sections, api = [], [], {}
    for document in sources[:2]:
        code, pending = False, None
        document_lines = (DOCS / document).read_text(encoding="utf-8-sig").splitlines()
        for number, raw in enumerate(document_lines, 1):
            if raw.startswith("```"):
                code = not code
                continue
            if raw.startswith("<!--"):
                continue
            anchor = re.fullmatch(r'<a id="([^"]+)"></a>', raw)
            if anchor:
                pending = anchor[1]
                continue
            if raw.startswith("|---"):
                continue
            style = "code" if code else "paragraph"
            text = raw
            if not code:
                heading = re.match(r"^(#{1,3}) (.*)$", raw)
                if heading:
                    style = "h" + str(len(heading[1]))
                    text = heading[2]
                    if style == "h1":
                        text = "Пошаговый гайд" if document == sources[0] else "Справочник методов и options/enum"
                    if style in ("h1", "h2"):
                        key = "section-" + str(len(sections) + 1)
                        sections.append({"title": text, "anchor": key, "index": len(rows), "style": style})
                text = re.sub(r"\[([^]]+)\]\(([^)]+)\)", r"\1 (\2)", text)
                text = text.replace("**", "").replace("`", "")
                if raw.startswith("|"):
                    parts = [part.strip().replace("\\|", "|") for part in re.split(r"(?<!\\)\|", raw.strip("|"))]
                    if len(parts) not in (2, 3):
                        raise ValueError("Guide table needs 2 or 3 actual columns: " + raw)
                    text = "\t".join(parts)
                    style = "table"
            row = {"text": text, "style": style, "source": document, "line": number}
            if style == "table":
                row["cells"] = parts
                row["tableHeader"] = number < len(document_lines) and document_lines[number].startswith("|---")
            if pending:
                row["anchor"] = pending
                api[pending] = len(rows)
                pending = None
            rows.append(row)
        if code:
            raise ValueError("Unclosed code fence: " + document)
    expected = {member["anchor"] for member in manifest["members"]}
    if set(api) != expected:
        raise ValueError("API Help anchors differ from actual manifest")
    # Все переходы используют адреса, имён на каждый метод не создаём.
    front = [{"text": "Добавление нового типа сечения", "style": "title"},
             {"text": "Оглавление. Переход раскрывает нужный подраздел через существующую навигацию Справки.", "style": "paragraph"}]
    for section in sections:
        front.append({"text": section["title"], "style": "toc", "link": section["anchor"]})
    front.append({"text": "", "style": "paragraph"})
    for section in sections:
        rows[section["index"]]["section"] = section["anchor"]
        section["index"] += len(front)
    for key in api:
        api[key] += len(front)
    rows = front + rows
    body_hash = hashlib.sha256("\n".join(row["text"] for row in rows).encode("utf-8")).hexdigest()
    rows.append({"text": "NDM.RECIPE.GUIDE/1|" + body_hash + "|" + sha(DOCS / sources[2]), "style": "footer"})
    return {"schema": 1, "title": "Добавление нового типа сечения", "bodySHA256": body_hash,
            "sourceHashes": {str((DOCS / source).relative_to(ROOT).as_posix()): sha(DOCS / source) for source in sources},
            "rows": rows, "sections": sections, "api": api,
            "apiStatus": manifest["status"], "members": len(expected)}


if __name__ == "__main__":
    result = build()
    (DOCS / "SECTION_RECIPE_GUIDE_LAYOUT.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": "GUIDE_SOURCE_SYNCHRONIZED", "rows": len(result["rows"]), "members": result["members"], "apiStatus": result["apiStatus"]}))
