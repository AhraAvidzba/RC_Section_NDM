Attribute VB_Name = "modTestGeometryReview"
Option Explicit
' ДЛЯ ТЕСТОВ: воспроизведение F01/F02 и проверка подписей отключённых граней.
Private mPassed As Long, mFailed As Long, mReport As String

Public Function RunGeometryReviewTests(Optional ByVal outputDirectory As String = "") As String
    mPassed = 0: mFailed = 0: mReport = ""
    TestPerPath 3, False: TestPerPath 3, True
    TestPerPath 15, False: TestPerPath 15, True
    TestRoundedLimit 1000#, outputDirectory: TestRoundedLimit 600#, outputDirectory
    TestRoundedLimit 600.001, outputDirectory
    TestDisabledSides "Rectangle", "rngRectSetGeometry", Array("H1.as_1", "H1.as_2", "B1.as_1", "B1.as_2")
    TestDisabledSides "RoundedSimple", "rngRoundedRectangleGeometry", Array("H.Left", "H.Right", "B.Top", "B.Bottom")
    TestDisabledSides "HollowCentered", "rngHollowRectangleGeometry", Array("H.Left", "H.Right", "B.Top", "B.Bottom")
    RunGeometryReviewTests = mReport & "TOTAL_GEOMETRY_REVIEW: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' Два независимых пути имеют одинаковый локальный отсчёт 1,3,...;
' смена порядка путей меняет публикацию, но не физические позиции ряда.
Private Sub TestPerPath(ByVal n As Long, ByVal reverse As Boolean)
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim builder As CSectionBuilder, model As CSectionModel, mesh As TSectionMeshOptions
    Dim i As Long, first As Long, second As Long, id As String, ids As Object, label As String
    label = "F01.n" & CStr(n) & ".reverse" & CStr(reverse)
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "ReviewPerPath"
    Set path = facade.CirclePath("Outer", 0#, 0#, 1000#): facade.Outer "OUTER", path.ID
    Set path = facade.LinePath("P1", 0#, 0#, 100#, 0#)
    Set path = facade.LinePath("P2", 0#, 200#, 100#, 200#)
    If reverse Then facade.PathGroup "Sides", Array("P2", "P1") Else facade.PathGroup "Sides", Array("P1", "P2")
    Set options = New CRebarRuleOptions
    facade.RebarByCount "Main", "Sides", n, 10#, options
    Set options = New CRebarRuleOptions: options.NormalOffset = 20#: options.EverySecond = True
    facade.RebarDependent "Extra", "Main", 8#, options
    mesh.StepX = 1000#: mesh.StepY = 1000#: mesh.BoundarySubdivisions = 1: mesh.MaterialID = 1
    Set builder = New CSectionBuilder: Set model = builder.Build(facade.Finish, mesh)
    Set ids = CreateObject("Scripting.Dictionary")
    For i = 1 To builder.Rules.PositionCount
        id = builder.Rules.PositionSourceID(i)
        Check label & ".unique." & CStr(i), Not ids.Exists(id)
        ids.Add id, True
        If builder.Rules.PositionRuleID(i) = "Extra" Then
            If builder.Rules.PositionPathID(i) = "P1" Then first = first + 1 Else second = second + 1
        End If
    Next i
    Check label & ".P1", first = (n + 1) \ 2
    Check label & ".P2", second = (n + 1) \ 2
    Check label & ".physical", model.RebarCount = 2 * n + n + 1
    mReport = mReport & "PER_PATH: " & label & "; extra=" & CStr(first) & "+" & CStr(second) & "; physical=" & CStr(model.RebarCount) & vbCrLf
    Exit Sub
Failed:
    Check label & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' Отключает по одной стороне через n=0 и d=0. Подписи остальных сторон
' обязаны совпасть с их фактическими стержнями; размеры сечения сохраняются.
Private Sub TestDisabledSides(ByVal fixture As String, ByVal tableName As String, ByVal groups As Variant)
    Dim config As Object, saved As Variant, table As Object, main As Long, extra As Long, side As Long, mode As Long, row As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim labels As Object, counts As Object, diameters As Object, group As Variant, i As Long, key As String, factor As Double
    Dim baseDimensions As Long, dimensions As Long, label As String
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    modTestShapeArchitecture.ConfigureShapeFixture "Circle": modTestShapeArchitecture.ConfigureShapeFixture fixture
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    Set table = ThisWorkbook.Names(tableName).RefersToRange
    main = HeaderRow(table, "as", 2): extra = HeaderRow(table, "d2", 2)
    For row = main + 1 To main + 4
        table.Cells(row, 2).Value2 = 50# / factor: table.Cells(row, 3).Value2 = 20# / factor
        table.Cells(row, 4).Value2 = 4: table.Cells(row, 5).Value2 = 0#: table.Cells(row, 6).Value2 = 0#
    Next row
    For row = extra + 1 To extra + 4: table.Cells(row, 2).Value2 = 0#: table.Cells(row, 5).Value2 = 0#: Next row
    Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
    For i = 1 To model.AnnotationCount
        If model.Annotations.AnnotationType(i) = "DIMENSION" Then baseDimensions = baseDimensions + 1
    Next i
    For side = 1 To 4
        For mode = 3 To 4
            label = "disabledSide." & fixture & "." & CStr(groups(side - 1)) & ".column" & CStr(mode)
            table.Cells(main + side, mode).Value2 = 0
            Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
            Set labels = CreateObject("Scripting.Dictionary"): Set counts = CreateObject("Scripting.Dictionary")
            Set diameters = CreateObject("Scripting.Dictionary")
            For i = 1 To model.RebarCount
                key = builder.Rebars.BarAnnotationGroupName(i)
                If Not counts.Exists(key) Then counts.Add key, 0: diameters.Add key, builder.Rebars.Diameter(i)
                counts(key) = CLng(counts(key)) + 1
            Next i
            dimensions = 0
            For i = 1 To model.AnnotationCount
                If model.Annotations.AnnotationType(i) = "REBAR_ANNOTATION" Then
                    key = Mid$(model.Annotations.AnnotationID(i), 7): labels.Add key, model.Annotations.Text(i)
                    Check label & ".noEmptyLabel." & key, counts.Exists(key)
                End If
                If model.Annotations.AnnotationType(i) = "DIMENSION" Then dimensions = dimensions + 1
            Next i
            Check label & ".off", Not labels.Exists(CStr(groups(side - 1))) And Not counts.Exists(CStr(groups(side - 1)))
            Check label & ".dimensions", dimensions = baseDimensions
            For Each group In counts.Keys
                Check label & ".otherLabel." & CStr(group), labels.Exists(CStr(group))
                If labels.Exists(CStr(group)) Then Check label & ".quantity." & CStr(group), _
                    CStr(labels(group)) = CStr(counts(group)) & ChrW$(216) & CStr(diameters(group))
            Next group
            table.Cells(main + side, 3).Value2 = 20# / factor: table.Cells(main + side, 4).Value2 = 4
        Next mode
    Next side
    GoTo Restore
Failed:
    Check label & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then config.Formula = saved
End Sub

' Прямые B исчезают при R1=B/2; боковой аналитический путь сохраняется.
' Ожидаемая длина — остаточная вертикаль плюс две четверти окружности R-as.
Private Sub TestRoundedLimit(ByVal height As Double, ByVal outputDirectory As String)
    Dim config As Object, saved As Variant, table As Object, header As Long, row As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder
    Dim model As CSectionModel, label As String, i As Long, expected As Double
    Dim definition As CSectionDefinition, paths As Variant, path As CGeometryPath, factor As Double, extra As Long
    Dim passed As Long, failed As Long, result As String
    On Error GoTo Failed
    label = "F02.H" & CStr(height)
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    modTestShapeArchitecture.ConfigureShapeFixture "RoundedSimple"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    Set table = ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersToRange
    table.Cells(3, 2).Value2 = 600# / factor: table.Cells(3, 3).Value2 = height / factor
    table.Cells(7, 2).Value2 = "Simple": table.Cells(7, 3).Value2 = "Simple"
    table.Cells(9, 2).Value2 = 300# / factor: table.Cells(9, 3).Value2 = 300# / factor
    header = HeaderRow(table, "as", 2): extra = HeaderRow(table, "d2", 2)
    For row = header + 1 To header + 4
        table.Cells(row, 2).Value2 = 50# / factor: table.Cells(row, 3).Value2 = 20# / factor
        ' В таблице сначала H.Left/H.Right, затем B.Top/B.Bottom.
        table.Cells(row, 4).Value2 = IIf(row < header + 3, 5, 0)
        table.Cells(row, 5).Value2 = 0#: table.Cells(row, 6).Value2 = 0#
    Next row
    For row = extra + 1 To extra + 4: table.Cells(row, 2).Value2 = 0#: table.Cells(row, 5).Value2 = 0#: Next row
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    expected = height - 600# + GEOM_PI * 250#
    For Each paths In Array("H.Left", "H.Right")
        Set path = definition.Path(CStr(paths))
        Check label & ".length." & CStr(paths), Abs(path.Length - expected) < 0.000001
    Next paths
    Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
    Check label & ".tenBars", model.RebarCount = 10
    Check label & ".contour", Abs(definition.ConcreteDomain.Area - (600# * (height - 600#) + GEOM_PI * 90000#)) < 0.00001
    For i = 1 To model.AnnotationCount
        Check label & ".noDisabledLabel." & CStr(i), model.Annotations.AnnotationID(i) <> "REBAR_B.Top" And model.Annotations.AnnotationID(i) <> "REBAR_B.Bottom"
    Next i
    result = modTestSectionRecipes.ReviewModelEquilibrium(label, model, settings, units, passed, failed)
    mPassed = mPassed + passed: mFailed = mFailed + failed: mReport = mReport & result
    If Len(outputDirectory) > 0 Then
        result = modTestShapeArchitecture.PreparedPreviewSnapshot(Replace$(label, ".", "_"), model, outputDirectory)
        Check label & ".preview", Len(result) > 0
    End If
    CheckRoundedWorkbookRun label, units, model
    GoTo Restore
Failed:
    Check label & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then config.Formula = saved
End Sub

' Запускает штатный workbook-конвейер одного LC и проверяет запись НДС.
' Это direct-state профиль; матрицы прочности/трещин здесь не запускаются.
Private Sub CheckRoundedWorkbookRun(ByVal label As String, ByVal units As CUnitSystem, ByVal model As CSectionModel)
    Dim loads As Object, profiles As Object, settings As Object, row As Long, key As String, message As String
    Dim properties As CSectionPropertiesCalculator
    Set properties = New CSectionPropertiesCalculator: properties.CalculateConcrete model
    Set settings = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    For row = 2 To settings.Rows.Count
        key = CStr(settings.Cells(row, 1).Value2)
        If key = "General.ExecutionReportEnabled" Or key = "Plot.AutoUpdate" Then settings.Cells(row, 2).Value2 = "No"
    Next row
    Set profiles = ThisWorkbook.Names("rngCalculationProfiles").RefersToRange
    For row = 2 To profiles.Rows.Count
        Select Case CStr(profiles.Cells(row, 2).Value2)
            Case "Calculation.Strength.DirectState": profiles.Cells(row, 3).Value2 = "Yes"
            Case "Calculation.Strength.Capacity", "Calculation.Crack.Width", "Calculation.Stability.Enabled", "Calculation.Crack.Longitudinal"
                profiles.Cells(row, 3).Value2 = "No"
        End Select
    Next row
    Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    loads.Offset(1, 0).Resize(loads.Rows.Count - 1, loads.Columns.Count).ClearContents
    loads.Cells(2, 1).Value2 = label: loads.Cells(2, 2).Value2 = -2# * properties.Area / units.InputForceToInternal(1#)
    loads.Cells(2, 3).Value2 = 0#: loads.Cells(2, 4).Value2 = 0#: loads.Cells(2, 5).Value2 = "PR1"
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    Check label & ".workbook.run", InStr(1, message, "ошиб", vbTextCompare) = 0 And InStr(1, message, "InputErr", vbTextCompare) = 0
    Check label & ".workbook.state", SnapshotAnchoredExtent(ThisWorkbook.Names("rngNDMElementResults").RefersToRange, False) > 1
    mReport = mReport & "WORKBOOK_RUN: " & label & "; " & message & vbCrLf
End Sub

' Находит строку настоящего заголовка fixture; перенос таблицы не меняет тест.
Private Function HeaderRow(ByVal table As Object, ByVal text As String, ByVal column As Long) As Long
    Dim row As Long
    For row = 1 To table.Rows.Count
        If CStr(table.Cells(row, column).Value2) = text Then HeaderRow = row: Exit Function
    Next row
    Err.Raise 5, "modTestGeometryReview", "Не найден заголовок " & text
End Function

Private Sub Check(ByVal name As String, ByVal success As Boolean)
    If success Then mPassed = mPassed + 1: mReport = mReport & "OK: " & name & vbCrLf Else mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
End Sub
