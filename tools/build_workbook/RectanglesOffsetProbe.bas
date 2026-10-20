' ДЛЯ ТЕСТОВ: добавляется только в modTestSectionRecipes отдельной копии.
Public Function DumpRectanglesOffsetBefore() As String
    DumpRectanglesOffsetBefore = CheckRectanglesOffsetCases(False)
End Function

Public Function RunRectanglesOffsetDimensionTests() As String
    RunRectanglesOffsetDimensionTests = CheckRectanglesOffsetCases(True)
End Function

' Проверяет величину размера и его реальные привязки, сохраняет всю расчётную геометрию для A/B.
Private Function CheckRectanglesOffsetCases(ByVal positiveDimension As Boolean) As String
    Dim config As Object, saved As Variant, inputRange As Object, offsetRow As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, reader As CSectionInputSnapshotReader
    Dim context As CSectionRecipeContext, builder As CSectionBuilder, model As CSectionModel
    Dim offset As Variant, i As Long, found As Long, output As String, expected As Double
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    modSolverWorkStats.ResetSectionEquilibriumSolveCount
    For Each offset In Array(-50#, 0#, 50#)
        config.Formula = saved
        modTestShapeArchitecture.ConfigureShapeFixture "Circle"
        modTestShapeArchitecture.ConfigureShapeFixture "TwoRectangles"
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set reader = New CSectionInputSnapshotReader
        Set context = reader.Capture(ThisWorkbook, Array("rngRectanglesGeometry"), units)
        offsetRow = context.Table("rngRectanglesGeometry").RowByKey("Rectangles.UpperOffsetX")
        Set inputRange = ThisWorkbook.Names("rngRectanglesGeometry").RefersToRange
        inputRange.Cells(offsetRow, 2).Value2 = CDbl(offset) / units.InputLengthToInternal(1#)
        ThisWorkbook.Worksheets("Config").Calculate
        settings.LoadFromWorkbook ThisWorkbook
        Set builder = New CSectionBuilder: Set model = BuildRecipe("Rectangles", settings, units, builder)
        Near "offset.input.preserved." & CStr(offset), units.InputLengthToInternal(CDbl(inputRange.Cells(offsetRow, 2).Value2)), CDbl(offset)
        found = 0
        For i = 1 To model.Annotations.Count
            If model.Annotations.AnnotationID(i) = "DIM_UPPER_OFFSET_X" Then
                found = found + 1
                expected = CDbl(offset): If positiveDimension Then expected = Abs(expected)
                Near "offset.dimension.value." & CStr(offset), model.Annotations.Value(i), expected
                Near "offset.dimension.left." & CStr(offset), GeomMin(model.Annotations.StartX(i), model.Annotations.EndX(i)), GeomMin(0#, CDbl(offset))
                Near "offset.dimension.right." & CStr(offset), GeomMax(model.Annotations.StartX(i), model.Annotations.EndX(i)), GeomMax(0#, CDbl(offset))
                Check "offset.dimension.label." & CStr(offset), model.Annotations.Text(i) = ChrW$(&H394)
                output = output & "OFFSET_CASE|offset=" & CStr(offset) & "|dimension=" & CStr(model.Annotations.Value(i)) & vbCrLf
            End If
        Next i
        Check "offset.dimension.presence." & CStr(offset), found = IIf(CDbl(offset) = 0#, 0, 1)
        If CDbl(offset) = 0# Then output = output & "OFFSET_CASE|offset=0|dimension=not displayed" & vbCrLf
        output = output & modTestShapeArchitecture.PreparedShapeSnapshot("Offset" & CStr(offset), model, builder.Mesh, builder.Rebars)
    Next offset
    Check "offset.no.equilibrium.solves", modSolverWorkStats.SectionEquilibriumSolveCount() = 0
    GoTo Finished
Failed:
    Check "offset.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    CheckRectanglesOffsetCases = output & mReport & "TOTAL_RECTANGLES_OFFSET: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function
