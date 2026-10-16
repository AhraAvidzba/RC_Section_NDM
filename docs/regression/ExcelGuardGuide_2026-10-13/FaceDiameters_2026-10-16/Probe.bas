' ДЛЯ ТЕСТОВ: дополнительный код только для отдельной проверочной копии книги.
' Рабочие рецепты, общий сборщик и исходная книга не изменяются.
Public Function DumpFaceDiameterProbe() As String
    Dim result As String, kind As Variant, mode As Long, permutation As Long
    Dim definition As CSectionDefinition, builder As CSectionBuilder, model As CSectionModel
    Dim mesh As TSectionMeshOptions, label As String, faces As Variant, dimensions As Variant
    Dim i As Long, annotations As CSectionAnnotations, names As Variant, name As Variant
    Dim table As CSectionInputTable, values As Variant, header As Long, column As Long
    On Error GoTo Failed
    ' Убираем пользовательские отступы от концов только в проверочной копии.
    For Each name In Array("rngRectSetGeometry", "rngRoundedRectangleGeometry", "rngHollowRectangleGeometry")
        values = ThisWorkbook.Names(CStr(name)).RefersToRange.Value2
        Set table = New CSectionInputTable: table.Initialize CStr(name), values
        header = table.RowByKey("as", 2)
        For Each kind In Array("StartSetback", "EndSetback")
            column = table.ColumnByCaption(header, CStr(kind))
            For i = 1 To IIf(CStr(name) = "rngRoundedRectangleGeometry", 4, 8)
                ThisWorkbook.Names(CStr(name)).RefersToRange.Cells(header + i, column).Value2 = 0#
            Next i
        Next kind
        For i = 1 To IIf(CStr(name) = "rngRoundedRectangleGeometry", 4, 8)
            result = result & "MAP" & vbTab & CStr(name) & vbTab & CStr(i - 1) & vbTab & CStr(values(header + i, 1)) & vbCrLf
        Next i
    Next name
    mesh.StepX = 200#: mesh.StepY = 200#: mesh.BoundarySubdivisions = 2: mesh.MaterialID = 1
    names = Array("Rectangle", "LSection", "TwoRectangles", "RoundedSimple", "RoundedTapered", _
        "RoundedMixed", "RoundedMixedLeft", "HollowCentered", "HollowOffset", "HollowSharp")
    For Each kind In names
        For mode = 0 To 2
            For permutation = 0 To 1
                label = CStr(kind) & ".m" & CStr(mode) & ".p" & CStr(permutation)
                result = result & "CASE" & vbTab & label & vbTab & CStr(kind) & vbTab & CStr(mode) & vbTab & CStr(permutation) & vbCrLf
                If Left$(CStr(kind), 7) = "Rounded" Then
                    dimensions = Array(1200#, 1600#, "Simple", "Simple", 250#, 300#, 180#, 200#, 160#, 170#)
                    If CStr(kind) = "RoundedTapered" Or CStr(kind) = "RoundedMixedLeft" Then dimensions(2) = "Tapered"
                    If CStr(kind) = "RoundedTapered" Or CStr(kind) = "RoundedMixed" Then dimensions(3) = "Tapered"
                    faces = FDFaces(4, mode, permutation)
                    Set definition = FixtureRoundedDefinition(dimensions, faces)
                ElseIf Left$(CStr(kind), 6) = "Hollow" Then
                    dimensions = Array(2000#, 2400#, 200#, 1000#, 1400#, 120#, 0#, 0#)
                    If CStr(kind) <> "HollowCentered" Then dimensions(6) = 125#: dimensions(7) = -125#
                    If CStr(kind) = "HollowSharp" Then dimensions(2) = 0#: dimensions(5) = 0#
                    faces = FDFaces(8, mode, permutation)
                    Set definition = FixtureHollowDefinition(dimensions, faces)
                Else
                    dimensions = Array(700#, 800#, 1200#, 500#, 0#, CStr(kind))
                    If CStr(kind) = "TwoRectangles" Then dimensions(4) = 175#
                    faces = FDRectFaces(mode, permutation)
                    Set definition = FDRectDefinition(dimensions, faces)
                End If
                result = result & FDValues("DIM", label, dimensions)
                For i = 0 To UBound(faces): result = result & FDValues("INPUT", label & "." & CStr(i), faces(i)): Next i
                Set builder = New CSectionBuilder: Set model = builder.Build(definition, mesh)
                For i = 1 To model.RebarCount
                    result = result & FDValues("BAR", label, Array(i, model.RebarX(i), model.RebarY(i), model.RebarDiameter(i), _
                        model.RebarArea(i), builder.Rebars.BarAnnotationGroupName(i), builder.Rebars.BarID(i)))
                Next i
                For i = 1 To builder.Rules.PositionCount
                    result = result & FDValues("POSITION", label, Array(builder.Rules.PositionPhysicalOwner(i), builder.Rules.PositionRuleID(i)))
                Next i
                Set annotations = model.Annotations
                For i = 1 To annotations.Count
                    If annotations.AnnotationType(i) = "REBAR_ANNOTATION" Then
                        result = result & FDValues("ANNOTATION", label, Array(annotations.AnnotationID(i), annotations.Text(i), _
                            annotations.StartX(i), annotations.StartY(i), annotations.EndX(i), annotations.EndY(i), _
                            annotations.NormalX(i), annotations.NormalY(i)))
                    End If
                Next i
                result = result & FDValues("BUILT", label, Array(model.RebarCount, builder.Rules.PositionCount, builder.Rules.MergedPositions, builder.Rules.Conflicts))
            Next permutation
        Next mode
    Next kind
    DumpFaceDiameterProbe = result & "TOTAL FACE DIAMETER BUILD: cases=60; failed=0" & vbCrLf
    Exit Function
Failed:
    DumpFaceDiameterProbe = result & "RUNTIME ERROR: " & label & "; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
End Function

' Заполняет RectSet по видимым названиям сторон, независимо от порядка строк.
Private Function FDRectDefinition(ByVal dimensions As Variant, ByVal faces As Variant) As CSectionDefinition
    Dim values As Variant, table As CSectionInputTable, recipe As ISectionRecipe
    Dim main As Long, extra As Long, i As Long, half As Long, row As Long, additionalRow As Long
    Dim caption As String, sides As Variant, pair As Variant
    values = ThisWorkbook.Names("rngRectSetGeometry").RefersToRange.Value2
    Set table = New CSectionInputTable: table.Initialize "rngRectSetGeometry", values
    values(table.RowByKey("RectSet.SectionType"), 2) = dimensions(5)
    values(table.RowByKey("RectSet.UpperOffsetX"), 2) = dimensions(4)
    row = table.RowByKey("H1") + 1
    values(row, 1) = dimensions(1): values(row, 2) = dimensions(0)
    values(row, 3) = dimensions(3): values(row, 4) = dimensions(2)
    main = table.RowByKey("as", 2): extra = table.RowByKey("d2", 2)
    sides = Array(" - левая", " - правая", " - верхняя", " - нижняя")
    For i = 0 To 3
        pair = faces(i)
        For half = 0 To 1
            caption = IIf(i < 2, "H", "B") & CStr(half + 1) & CStr(sides(i))
            row = table.RowByAnyKey(Array(caption), 1, main + 1, main + 8)
            additionalRow = table.RowByAnyKey(Array(caption), 1, extra + 1, extra + 8)
            values(row, 2) = pair(half): values(row, 3) = pair(2 + half): values(row, 4) = pair(4 + half)
            values(row, 5) = pair(6 + 2 * half): values(row, 6) = pair(7 + 2 * half)
            values(additionalRow, 2) = pair(10 + half): values(additionalRow, 5) = pair(12 + half)
            values(additionalRow, 3) = pair(14): values(additionalRow, 6) = pair(15)
            values(additionalRow, 4) = pair(16): values(additionalRow, 7) = pair(17)
        Next half
    Next i
    Set recipe = New CSectionRecipeRectSet
    Set FDRectDefinition = FixtureRecipeDefinition(recipe, "rngRectSetGeometry", values)
End Function

' Возвращает разные диаметры для каждой стороны; дополнительные ряды также отличаются.
Private Function FDFaces(ByVal count As Long, ByVal mode As Long, ByVal permutation As Long) As Variant
    Dim faces() As Variant, d1 As Variant, d2 As Variant, d3 As Variant, i As Long, index As Long, n As Long
    d1 = Array(12#, 16#, 20#, 25#, 28#, 32#, 36#, 40#)
    d2 = Array(10#, 12#, 16#, 20#, 25#, 28#, 32#, 36#)
    d3 = Array(8#, 10#, 12#, 16#, 20#, 25#, 28#, 32#)
    ReDim faces(0 To count - 1)
    For i = 0 To count - 1
        index = i: If permutation = 1 Then index = count - 1 - i
        n = 5: If i Mod 4 >= 2 Then n = 9
        If count = 8 Then n = 9
        faces(i) = Array(90# + 7# * i, d1(index), n, IIf(mode > 0, d2(index), 0#), IIf(mode > 0, d3(index), 0#), _
            IIf(mode = 2, "SideBySide", "Stacked"), IIf(mode = 2, "SideBySide", "Stacked"), "EachBar", IIf(mode = 2, "EverySecondBar", "EachBar"))
    Next i
    FDFaces = faces
End Function

' Возвращает параметры обеих частей RectSet в порядке строк таблицы Config.
Private Function FDRectFaces(ByVal mode As Long, ByVal permutation As Long) As Variant
    Dim faces(0 To 3) As Variant, d1 As Variant, d2 As Variant, d3 As Variant, i As Long, a As Long, b As Long
    d1 = Array(12#, 16#, 20#, 25#, 28#, 32#, 36#, 40#)
    d2 = Array(10#, 12#, 16#, 20#, 25#, 28#, 32#, 36#)
    d3 = Array(8#, 10#, 12#, 16#, 20#, 25#, 28#, 32#)
    For i = 0 To 3
        a = i * 2: b = a + 1
        If permutation = 1 Then a = 7 - a: b = 7 - b
        faces(i) = Array(70# + 7# * i, 75# + 7# * i, d1(a), d1(b), 5 + i, 6 + i, 100#, 100#, 100#, 100#, _
            IIf(mode > 0, d2(a), 0#), IIf(mode > 0, d2(b), 0#), IIf(mode > 0, d3(a), 0#), IIf(mode > 0, d3(b), 0#), _
            IIf(mode = 2, "SideBySide", "Stacked"), IIf(mode = 2, "SideBySide", "Stacked"), "EachBar", IIf(mode = 2, "EverySecondBar", "EachBar"))
    Next i
    FDRectFaces = faces
End Function

' Сохраняет числа с достаточной точностью для независимой проверки координат в Python.
Private Function FDValues(ByVal kind As String, ByVal label As String, ByVal values As Variant) As String
    Dim value As Variant
    FDValues = kind & vbTab & label
    For Each value In values
        If IsNumeric(value) And VarType(value) <> vbString Then
            FDValues = FDValues & vbTab & Replace$(Format$(CDbl(value), "0.000000000000"), ",", ".")
        Else
            FDValues = FDValues & vbTab & CStr(value)
        End If
    Next value
    FDValues = FDValues & vbCrLf
End Function
