Attribute VB_Name = "modTestGeometryReview"
Option Explicit
' ДЛЯ ТЕСТОВ: воспроизведение F01/F02 и проверка подписей отключённых граней.
Private mPassed As Long, mFailed As Long, mReport As String

Public Function RunGeometryReviewTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    TestPerPath 3, False: TestPerPath 3, True
    TestPerPath 15, False: TestPerPath 15, True
    TestRoundedLimit 1000#: TestRoundedLimit 600#
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

' Прямые B исчезают при R1=B/2; боковой аналитический путь сохраняется.
' Ожидаемая длина — остаточная вертикаль плюс две четверти окружности R-as.
Private Sub TestRoundedLimit(ByVal height As Double)
    Dim config As Object, saved As Variant, table As Object, header As Long, row As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder
    Dim model As CSectionModel, label As String, i As Long, expected As Double
    Dim definition As CSectionDefinition, paths As Variant, path As CGeometryPath, factor As Double, extra As Long
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
    GoTo Restore
Failed:
    Check label & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then config.Formula = saved
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
