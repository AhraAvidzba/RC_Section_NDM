"""ДЛЯ ТЕСТОВ: готовит изолированный guard/test diff без изменения frozen source."""
from pathlib import Path
import difflib
import hashlib
import json

ROOT = Path(__file__).resolve().parents[3]
EVIDENCE = Path(__file__).resolve().parent
OUTPUT = EVIDENCE / 'DeferredContourSnapshotFix01'


def digest(value):
    return hashlib.sha256(value).hexdigest().upper()


def prepare():
    if OUTPUT.exists():
        raise ValueError('Use a fresh isolated patch directory')
    sources = ['tests/modTestAutoCADContours.bas', 'src/Excel/modSectionSnapshot.bas']
    original = {name: (ROOT / name).read_bytes() for name in sources}
    texts = {name: data.decode('utf-8-sig').replace('\r\n', '\n') for name, data in original.items()}
    test = texts[sources[0]]

    def replace_once(text, before, after):
        if text.count(before) != 1:
            raise ValueError('Expected exactly one source block: ' + repr(before[:100]))
        return text.replace(before, after)

    test = replace_once(test,
        '    Dim settings As CSystemSettingsReader, units As CUnitSystem, unitSheet As Object\n',
        '    Dim settings As CSystemSettingsReader, units As CUnitSystem, unitSheet As Object\n'
        '    Dim geometryAnchor As Object, contourAnchor As Object, geometryData As Variant, geometryRows As Long, geometryColumns As Long\n')
    test = replace_once(test,
        '    Set marker = sheet.Cells(22, 46): marker.Value2 = "USER-GAP"\n'
        '    TestCheck "postAudit03.contours.geometry15", sheet.Cells(20, 27).Value2 = vbNullString\n',
        '    Set geometryAnchor = fixture.Names.Item("rngNDMSectionGeometry").RefersToRange\n'
        '    Set contourAnchor = fixture.Names.Item("rngNDMSectionContours").RefersToRange\n'
        '    geometryColumns = writer.OutputColumnCount("rngNDMSectionGeometry"): geometryRows = model.ConcreteCount + model.RebarCount + 1\n'
        '    geometryData = geometryAnchor.Resize(geometryRows, geometryColumns).Value2\n'
        '    Set marker = sheet.Cells(geometryAnchor.Row + 2, contourAnchor.Column + writer.OutputColumnCount("rngNDMSectionContours") + 1)\n'
        '    marker.Value2 = "USER-GAP"\n'
        '    TestCheck "postAudit03.contours.markerOutside17Columns", writer.OutputColumnCount("rngNDMSectionContours") = 17 And _\n'
        '        marker.Column > contourAnchor.Column + 16 And marker.Column < fixture.Names.Item("rngNDMSectionProperties").RefersToRange.Column\n'
        '    TestCheck "postAudit03.contours.geometry16", geometryColumns = 16 And AnchoredColumnCount(geometryAnchor) = 16 And _\n'
        '        Len(CStr(geometryAnchor.Offset(0, geometryColumns).Value2)) = 0\n'
        '    TestCheck "postAudit03.contours.geometry16.localIxyHeader", ResultColumn(geometryData, "LocalIxy") = 14\n'
        '    Set restored = ReadSectionGeometryFromResults(fixture)\n'
        '    TestCheckArea "postAudit03.contours.geometry16.localIxyRoundTrip", restored.ConcreteLocalIxy(1), 12#\n'
        '    geometryAnchor.Offset(0, geometryColumns - 1).Resize(geometryRows, 1).ClearContents\n'
        '    TestCheck "postAudit03.contours.legacy15.width", AnchoredColumnCount(geometryAnchor) = 15\n'
        '    Set restored = ReadSectionGeometryFromResults(fixture, "Legacy15", True)\n'
        '    TestCheckArea "postAudit03.contours.legacy15.area", restored.ConcreteArea(1), 10000#\n'
        '    TestCheckArea "postAudit03.contours.legacy15.localIx", restored.ConcreteLocalIx(1), 123456#\n'
        '    TestCheckArea "postAudit03.contours.legacy15.localIy", restored.ConcreteLocalIy(1), 654321#\n'
        '    TestCheckArea "postAudit03.contours.legacy15.localIxy", restored.ConcreteLocalIxy(1), 12#\n'
        '    geometryAnchor.Resize(geometryRows, geometryColumns).Value2 = geometryData\n')
    test = replace_once(test,
        '        TestCheckArea "postAudit03.contours.mechanicalIx." & CStr(unitText), restored.ConcreteLocalIx(1), 123456#\n',
        '        TestCheckArea "postAudit03.contours.mechanicalIx." & CStr(unitText), restored.ConcreteLocalIx(1), 123456#\n'
        '        TestCheckArea "postAudit03.contours.mechanicalIxy." & CStr(unitText), restored.ConcreteLocalIxy(1), 12#\n')
    test = replace_once(test,
        '    sheet.Cells(20, 27).Value2 = "EndX, mm"\n'
        '    On Error Resume Next\n'
        '    Set restored = ReadSectionGeometryFromResults(fixture): number = Err.Number: Err.Clear\n'
        '    On Error GoTo Failed\n'
        '    TestCheck "postAudit03.contours.expandedGeometryRejected", number = vbObjectError + 4355\n',
        '    fixture.Names.Add "rngNDMSectionContours", "=Results!$GR$20"\n'
        '    Set readBack = New CSectionContours: writer.WriteSectionContours fixture, readBack, Empty\n'
        '    geometryAnchor.Offset(0, geometryColumns).Value2 = "UnexpectedGeometryField"\n'
        '    TestCheck "postAudit03.contours.expandedGeometryIs17", AnchoredColumnCount(geometryAnchor) = 17\n'
        '    On Error Resume Next\n'
        '    Set restored = ReadSectionGeometryFromResults(fixture): number = Err.Number: description = Err.Description: Err.Clear\n'
        '    On Error GoTo Failed\n'
        '    TestCheck "postAudit03.contours.expandedGeometryRejected", number = vbObjectError + 4355\n'
        '    TestCheck "postAudit03.contours.expandedGeometryAddress", InStr(description, geometryAnchor.Resize(1, 17).Address(False, False)) > 0\n'
        '    geometryAnchor.Offset(0, geometryColumns).ClearContents\n'
        '    Set restored = ReadSectionGeometryFromResults(fixture)\n'
        '    TestCheckArea "postAudit03.contours.geometry16.restoredIxy", restored.ConcreteLocalIxy(1), 12#\n')
    reader = replace_once(texts[sources[1]],
        '    If UBound(data, 2) < 14 Then Err.Raise vbObjectError + 4355, "ReadSectionGeometryFromResults", _\n'
        '        "Неверная структура rngNDMSectionGeometry. Нажмите «Сформировать сечение»."\n',
        '    If UBound(data, 2) < 14 Or UBound(data, 2) > 16 Then Err.Raise vbObjectError + 4355, "ReadSectionGeometryFromResults", _\n'
        '        "Неверная структура rngNDMSectionGeometry: " & anchor.Parent.Name & "!" & _\n'
        '        anchor.Resize(1, UBound(data, 2)).Address(False, False) & ". Допустимы 14 обязательных полей и только Comment/ExtraData после них. Нажмите «Сформировать сечение»."\n')
    revised = {sources[0]: test, sources[1]: reader}
    OUTPUT.mkdir(parents=True)
    records, patch = [], []
    for relative in sources:
        data = original[relative]
        newline = '\r\n' if b'\r\n' in data else '\n'
        encoded = revised[relative].replace('\n', newline).encode('utf-8')
        if data.startswith(b'\xef\xbb\xbf'):
            encoded = b'\xef\xbb\xbf' + encoded
        destination = OUTPUT / 'Candidate' / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(encoded)
        records.append({'Source': relative, 'OriginalSHA256': digest(data), 'Candidate': str(destination.relative_to(ROOT)),
                        'CandidateSHA256': digest(encoded), 'SourceUnchanged': (ROOT / relative).read_bytes() == data})
        patch.extend(difflib.unified_diff(texts[relative].splitlines(True), revised[relative].splitlines(True),
                                         fromfile='a/' + relative, tofile='b/' + relative))
    (OUTPUT / 'DeferredPatch.diff').write_text(''.join(patch), encoding='utf-8')
    primary = EVIDENCE / 'FinalViews02_No.txt'
    raw = primary.read_bytes()
    text = raw.decode('utf-16' if raw.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig')
    lines = text.splitlines()
    manifest = {'Status': 'PREPARED_NOT_APPLIED_NATIVE_CONFIRMATION_PENDING', 'Files': records,
                'Primary': str(primary.relative_to(ROOT)), 'PrimarySHA256': digest(raw), 'PrimaryLinesRead': len(lines),
                'AllFailures': [line for line in lines if line.startswith('FAIL:')],
                'PrimaryTotals': [line for line in lines if line.startswith('TOTAL')],
                'Cause': ['Common geometry16 and contours17 already existed in HEAD; this task did not introduce their widths.',
                          'Old geometry15 checks and expanded fixture wrote inside legitimate16th ExtraData column.',
                          'Old marker column46 is the last column of contour block30..46, not a gap; moved by named anchor/output width into48, before properties49.',
                          'Reader lower bound14 has no upper guard; strict unknown17 rejection demanded by the old test was never truly exercised. New guard rejects>16 before consuming or mutating data.'],
                'DirectedPlan': ['Apply both exact source hashes only after all frozen processes stop; compile a fresh epoch.',
                                 'Run modTestAutoCADContours.RunPostAudit03ContourSnapshotTests in No and Yes sequentially; preserve original5FAIL primary.',
                                 'Verify new16 LocalIxy roundtrip, explicit legacy15 without ExtraData, mm/cm/m Ixy, gap marker, width17 rejection/error address and restored16 acceptance.',
                                 'Repeat affected snapshot reader gates for mandatory Contour+legacy before publication; regenerate actual AllCode from final book.'],
                'FullRegressionPass': False, 'ExcelStarted': False, 'WorkingSourcesEdited': False}
    (OUTPUT / 'DeferredSourceManifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'Status': manifest['Status'], 'Files': records, 'Failures': manifest['AllFailures']}, ensure_ascii=True))


if __name__ == '__main__':
    prepare()
