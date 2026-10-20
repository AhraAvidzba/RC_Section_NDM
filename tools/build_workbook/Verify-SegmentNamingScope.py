"""Проверяет границы изменения исходников: новые компоненты и посторонняя математика запрещены."""
import json
import re
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
D = ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/SegmentNaming_2026-10-16'
ALLOWED = {
    'CPolygonSpec': {'AddVertex', 'OutgoingEdgeID', 'CornerID', 'CopyKernelData'},
    'CGeometryPath': {'AppendInterval', 'CopyCollapsedAnchors'},
    'CSectionGeometryQuery': {'RoundedPolygon', 'RoundedPolygonKernel', 'OffsetRoundedPolygon',
                              'OffsetRoundedPolygonKernel', 'PolygonKernelData', 'OffsetPath'},
    'CSectionDefinitionBuilder': {'RectanglePath'},
    'CSectionRecipePointedSection': {'DescribePolygon'},
    'CSectionRecipeHollowRectangle': {'DescribePolygon'},
}


def code(body):
    lines = []
    for line in body.splitlines():
        out, quoted, i = '', False, 0
        while i < len(line):
            ch = line[i]
            if ch == '"':
                quoted = not quoted
            if ch == "'" and not quoted:
                break
            out += ch if quoted else ch.lower()
            i += 1
        if out.strip():
            lines.append(out.strip())
    return '\n'.join(lines)


def procedures(body):
    pattern = r'(?m)^(?:Public |Private |Friend )?(?:Static )?(Sub|Function|Property Get|Property Let|Property Set) (\w+)\b'
    matches = list(re.finditer(pattern, body))
    blocks = {}
    for i, match in enumerate(matches):
        blocks[match[2]] = body[match.start():matches[i+1].start() if i+1 < len(matches) else len(body)]
    declarations = body[:matches[0].start()] if matches else body
    return code(declarations), {name:code(value) for name,value in blocks.items()}


def verify():
    current = {path.relative_to(ROOT).as_posix():path for path in (ROOT/'src').rglob('*') if path.suffix in ('.bas','.cls')}
    changed, errors = {}, []
    with ZipFile(D/'BaselineSource.zip') as archive:
        old_paths = {name for name in archive.namelist() if name.startswith('src/') and name.endswith(('.bas','.cls'))}
        assert old_paths == current.keys(), 'Source component inventory changed'
        for name,path in current.items():
            old = archive.read(name).decode('utf-8-sig').replace('\r\n','\n')
            new = path.read_text(encoding='utf-8-sig')
            if code(old) == code(new):
                continue
            owner = path.stem
            if owner not in ALLOWED:
                errors.append('Unexpected source change: '+name)
                continue
            old_header, a = procedures(old)
            new_header, b = procedures(new)
            if old_header != new_header and owner != 'CPolygonSpec':
                errors.append('Class fields changed: '+name)
            methods = {key for key in a.keys() | b.keys() if a.get(key) != b.get(key)}
            if not methods <= ALLOWED[owner]:
                errors.append('Unexpected procedure change: '+name+': '+repr(methods))
            changed[owner] = sorted(methods)
    result = {'Passed':not errors,'SourceComponents':len(current),'NewComponents':0,
              'ChangedRuntimeProcedures':changed,'UnchangedOtherSourceCode':True,'Errors':errors}
    (D/'SourceScope.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False))
    assert not errors


if __name__ == '__main__':
    verify()
