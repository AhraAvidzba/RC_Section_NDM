' ==========================================================================
' ДЛЯ ТЕСТОВ — вспомогательный фрагмент частной приёмки; в выпуск не включается.
' ==========================================================================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет исходный знаковый размер смещения прямоугольников для воспроизводимого сравнения до изменения.
Public Function TestDumpRectanglesOffsetBefore() As String
    TestDumpRectanglesOffsetBefore = TestCheckRectanglesOffsetCases(False)
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет положительную величину размера смещения и реальные привязки, сохраняя геометрию для сравнения.
Public Function RunRectanglesOffsetDimensionTests() As String
    RunRectanglesOffsetDimensionTests = TestCheckRectanglesOffsetCases(True)
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет смещения -50, 0 и 50, значение и привязки аннотации, отсутствие нового solve и восстановление таблицы.
Private Function TestCheckRectanglesOffsetCases(ByVal positiveDimension As Boolean) As String
    Dim config As Object, saved As Variant, inputRange As Object, offsetCell As Object
    Dim settings As CSystemSettingsReader, units As CUnitSystem, reader As CSectionInputSnapshotReader
    Dim context As CSectionRecipeContext, builder As CSectionBuilder, model As CSectionModel
    Dim offset As Variant, i As Long, found As Long, output As String, expected As Double
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    modSolverWorkStats.ResetSectionEquilibriumSolveCount
    For Each offset In Array(-50#, 0#, 50#)
        TestFixtureRestoreInputs config, saved
        modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
        modTestShapeArchitecture.TestConfigureShapeFixture "TwoRectangles"
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set reader = New CSectionInputSnapshotReader
        Set context = reader.Capture(ThisWorkbook, Array("rngRectanglesGeometry"), units)
        Set inputRange = ThisWorkbook.Names("rngRectanglesGeometry").RefersToRange
        Set offsetCell = TestFixtureInputCell(inputRange, "Detail", "XOffset", "val") ' Текущая таблица хранит поле по блоку и строке, независимо от положения на листе.
        offsetCell.Value2 = CDbl(offset) / units.InputLengthToInternal(1#)
        ThisWorkbook.Worksheets("Config").Calculate
        settings.LoadFromWorkbook ThisWorkbook
        Set builder = New CSectionBuilder: Set model = TestBuildRecipe("Rectangles", settings, units, builder)
        TestNear "offset.input.preserved." & CStr(offset), units.InputLengthToInternal(CDbl(offsetCell.Value2)), CDbl(offset)
        found = 0
        For i = 1 To model.Annotations.Count
            If model.Annotations.AnnotationID(i) = "DIM_UPPER_OFFSET_X" Then
                found = found + 1
                expected = CDbl(offset): If positiveDimension Then expected = Abs(expected)
                TestNear "offset.dimension.value." & CStr(offset), model.Annotations.Value(i), expected
                TestNear "offset.dimension.left." & CStr(offset), GeomMin(model.Annotations.StartX(i), model.Annotations.EndX(i)), GeomMin(0#, CDbl(offset))
                TestNear "offset.dimension.right." & CStr(offset), GeomMax(model.Annotations.StartX(i), model.Annotations.EndX(i)), GeomMax(0#, CDbl(offset))
                TestCheck "offset.dimension.label." & CStr(offset), model.Annotations.Text(i) = ChrW$(&H394)
                output = output & "OFFSET_CASE|offset=" & CStr(offset) & "|dimension=" & CStr(model.Annotations.Value(i)) & vbCrLf
            End If
        Next i
        TestCheck "offset.dimension.presence." & CStr(offset), found = IIf(CDbl(offset) = 0#, 0, 1)
        If CDbl(offset) = 0# Then output = output & "OFFSET_CASE|offset=0|dimension=not displayed" & vbCrLf
        output = output & modTestShapeArchitecture.TestPreparedShapeSnapshot("Offset" & CStr(offset), model, builder.Mesh, builder.Rebars)
    Next offset
    TestCheck "offset.no.equilibrium.solves", modSolverWorkStats.SectionEquilibriumSolveCount() = 0
    GoTo Finished
Failed:
    TestCheck "offset.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    TestCheckRectanglesOffsetCases = output & mReport & "TOTAL_RECTANGLES_OFFSET: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function
