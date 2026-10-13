"""Краткий гайд на листе; API автора сечения находится в модуле комментариев VBE."""
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
    source = DOCS / 'ADDING_SECTION_TYPE.md'
    rows = []
    code = False
    for number, line in enumerate(source.read_text(encoding='utf-8-sig').splitlines(), 1):
        if line.startswith('```'):
            code = not code
            continue
        style = 'code' if code else 'paragraph'
        heading = re.match(r'^(#{1,2}) (.*)$', line)
        text = line
        if not code and heading:
            style = 'title' if len(heading[1]) == 1 else 'h2'
            text = heading[2]
        rows.append({'text': text, 'style': style, 'source': 'ADDING_SECTION_TYPE.md', 'line': number})
    body_hash = hashlib.sha256('\n'.join(row['text'] for row in rows).encode('utf-8')).hexdigest()
    rows.append({'text': 'NDM.RECIPE.GUIDE/2|'+body_hash+'|'+sha(ROOT/'src/Common/modSectionRecipeApiGuide.bas'), 'style': 'footer'})
    return {'schema': 2, 'title': 'Добавление нового типа сечения', 'bodySHA256': body_hash,
            'sourceHashes': {'docs/geometry_refactor/ADDING_SECTION_TYPE.md': sha(source),
                             'src/Common/modSectionRecipeApiGuide.bas': sha(ROOT/'src/Common/modSectionRecipeApiGuide.bas')},
            'rows': rows, 'sections': [], 'api': {}, 'apiStatus': 'COMMENT_ONLY_VBE_API',
            'members': len(json.loads((DOCS/'SECTION_RECIPE_AUTHOR_MANIFEST.json').read_text(encoding='utf-8-sig'))['members'])}


if __name__ == "__main__":
    result = build()
    (DOCS / "SECTION_RECIPE_GUIDE_LAYOUT.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": "GUIDE_SOURCE_SYNCHRONIZED", "rows": len(result["rows"]), "members": result["members"], "apiStatus": result["apiStatus"]}))
