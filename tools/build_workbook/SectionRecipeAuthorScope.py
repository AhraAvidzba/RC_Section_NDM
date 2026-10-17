"""Разделы справочника автора сечения; внутренний API остаётся в полном каталоге."""

# Имена нужны только для отбора и порядка. Сигнатуры, параметры и значения
# по умолчанию всегда читаются из действующих объявлений VBA.
SECTIONS = [
    ('01', 'Начало и завершение описания', [
        ('ISectionRecipe', 'CreateDefinition'),
        ('CSectionDefinitionBuilder', 'Initialize Finish'),
        ('CSectionTypeRegistry', 'RegisterRecipe'),
    ]),
    ('02', 'Входные размеры и единицы', [
        ('CSectionRecipeContext', 'Reader Table Units'),
        ('CSectionInputReader', '*'),
        ('CSectionInputTable', 'Name RowCount ColumnCount SourceReference CellValue IsBlank Text Number Count Choice Length PositiveLength NonNegativeLength OptionalDiameter OptionalRowOffset ColumnByCaption RowByKey RowByAnyKey'),
        ('CUnitSystem', 'InputLengthToInternal InputAreaToInternal'),
    ]),
    ('03', 'Границы бетона и отверстий', [
        ('CPolygonSpec', '*'),
        ('modSectionRecipeTypes', 'EContourNamingPolicy'),
        ('CSectionDefinitionBuilder', 'CirclePath LinePath ArcPath RectanglePath RoundedPolygon AdjacentRectangles AddPath Outer Opening'),
    ]),
    ('04', 'Линии центров стержней и преобразования', [
        ('CSectionDefinitionBuilder', 'Offset OffsetRoundedPolygon ConstrainLine Chain Slice Transform PathGroup'),
    ]),
    ('05', 'Размещение арматуры', [
        ('CSectionDefinitionBuilder', 'RebarByCount RebarByMaxStep RebarByFixedStep RebarAtStations RebarDependent RebarProjected'),
    ]),
    ('06', 'Условия размещения стержней', [
        ('CRebarRuleOptions', '*'),
        ('CRebarContactBuilder', 'Initialize AppendRow'),
        ('modSectionRecipeTypes', 'ERebarOffsetAxis ERebarEndMode ERebarDistributionScope ERemainderPlacement ERebarPublication ERebarPositionOrder ERebarCoincidence ERebarPlacementCheck ERebarOverlapPolicy'),
    ]),
    ('07', 'Размеры, пояснения и подписи арматуры', [
        ('CSectionDefinitionBuilder', 'Anchor PathAnchor TransformAnchor VerticalExtent HorizontalExtent HorizontalDistance VerticalDistance Dimension PathLength Radius Diameter Text RebarLabel RebarRowLabel RebarRowExtentLabel RebarRowExtremeLabel'),
    ]),
    ('08', 'Составные линии и отдельные участки', [
        ('CGeometryPath', 'Initialize AddLine AddArc Seal ID SegmentCount IsClosed Length SegmentID SegmentIndex AnchorStation ResolveAnchor BetweenAnchors AtDistance Slice Reversed Transformed WithID'),
        ('modGeometryTypes', 'GEOM_PI TPoint2D'),
    ]),
    ('09', 'Адреса ввода и сообщения об ошибках', [
        ('CSectionDefinitionBuilder', 'Source LastGeometryDiagnostic'),
        ('CGeometryPath', 'LastDiagnostic'),
        ('CGeometryDiagnostic', '*'),
    ]),
]


def selected_members(inventory):
    index = {(m['owner'], m['name']): m for m in inventory['members']}
    result, seen = [], set()
    for number, title, groups in SECTIONS:
        members = []
        for owner, names in groups:
            chosen = ([m for m in inventory['members'] if m['owner'] == owner]
                      if names == '*' else [index[(owner, name)] for name in names.split()])
            assert chosen, owner
            for member in chosen:
                assert member['id'] not in seen, member['id']
                seen.add(member['id'])
                members.append(member)
        result.append((number, title, members))
    # Любое расширение декларативного фасада должно получить место в справочнике.
    assert {m['id'] for m in inventory['members'] if m['owner'] == 'CSectionDefinitionBuilder'} <= seen
    return result
