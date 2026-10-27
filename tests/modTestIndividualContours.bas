Attribute VB_Name = "modTestIndividualContours"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: индивидуальные контуры связанных рядов без переноса стержней
' ==========================================================================
Private mPassed As Long, mFailed As Long ' Число успешных проверок и отказов текущего набора.
Private mReport As String ' Фактические координаты, пропуски и результаты проверок.
Private Type TPerformanceReal
    Value As Double ' Исходное число без округления для побитного сравнения версий.
End Type
Private Type TPerformanceBytes
    Value(0 To 7) As Byte ' Восемь исходных байтов Double.
End Type

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отдельный верх, все стороны, нечётные позиции и обращённые линии.
Public Function RunIndividualContourTests() As String
    Dim allSides As Long, reverse As Long, everySecond As Long, location As Long, disabled As Long
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    For allSides = 0 To 1
        For reverse = 0 To 1
            For everySecond = 0 To 1
                TestRectangle allSides = 1, reverse = 1, everySecond = 1
            Next everySecond
        Next reverse
    Next allSides
    TestRoundedCorners False
    TestRoundedCorners True
    TestMidplaneMetadata -1#
    TestMidplaneMetadata 1#
    TestIndependentContourDomains
    For location = 0 To 1
        For everySecond = 0 To 1
            For disabled = 0 To 1
                TestContactVariants location = 1, everySecond = 1, disabled = 1
            Next disabled
        Next everySecond
    Next location
    GoTo Finished
Failed:
    TestCheck "individual.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunIndividualContourTests = mReport & "TOTAL_INDIVIDUAL_CONTOURS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' время всей сборки геометрии и рядов при 100/1000/5000 основных стержней.
' Измерение заканчивается до проверок и вывода; сетка и поиск НДС сюда не входят.
Public Function RunIndividualContourPerformanceTests() As String
    Dim n As Variant, repeat As Long, start As Double, seconds As Double, facade As CSectionDefinitionBuilder, path As CGeometryPath
    Dim options As CRebarRuleOptions, definition As CSectionDefinition, executor As CRebarRuleExecutor, layout As CRebarLayout
    Dim i As Long, owner As Long, parent As TRebarPosition, rule As TRebarRule, lines() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    For Each n In Array(100, 1000, 5000)
        For repeat = 1 To 3
            start = Timer
            Set facade = New CSectionDefinitionBuilder: facade.Initialize "ContourPerformance"
            Set path = facade.RectanglePath("Boundary", 0#, 0#, CDbl(n) * 40# + 100#, 1000#): facade.Outer "Outer", path.ID
            Set path = facade.LinePath("Axis", 50#, 950#, CDbl(n) * 40# + 50#, 950#)
            Set options = New CRebarRuleOptions: options.SourceReference = "ContourPerformance[Main]"
            facade.RebarByCount "Main", path.ID, CLng(n), 10#, options
            facade.RebarContour "Main", "Outer", "Boundary", 50#, "R2", "R3", normalSign:=-1#
            Set options = New CRebarRuleOptions: options.ValidateRowOffset = True: options.NormalOffset = -100#
            options.SourceReference = "ContourPerformance[R2]": facade.RebarDependent "R2", "Main", 8#, options
            Set options = New CRebarRuleOptions: options.ValidateRowOffset = True: options.NormalOffset = -200#
            options.SourceReference = "ContourPerformance[R3]": facade.RebarDependent "R3", "Main", 6#, options
            Set definition = facade.Finish: Set executor = New CRebarRuleExecutor: Set layout = TestBuildLayout(definition, executor)
            seconds = Timer - start: If seconds < 0# Then seconds = seconds + 86400#
            mReport = mReport & "PERF_ROW: n=" & CStr(n) & "; repeat=" & CStr(repeat) & "; seconds=" & CStr(seconds) & "; physical=" & CStr(layout.Count) & vbCrLf
            TestCheck "performance.physical.count", layout.Count = 3 * CLng(n)
            TestCheck "performance.no.merge", executor.MergedPositions = 0
            TestCheck "performance.no.skip", Not TestHasOffsetSkip(executor)
            For i = 1 To executor.PositionCount
                If executor.PositionRuleID(i) <> "Main" Then
                    executor.TestCopyTransfer i, parent, rule: owner = executor.PositionPhysicalOwner(i)
                    TestNear "performance.preserved.x", layout.X(owner), parent.X - rule.NormalOffset * parent.TangentY
                    TestNear "performance.preserved.y", layout.Y(owner), parent.Y + rule.NormalOffset * parent.TangentX
                End If
            Next i
            If repeat = 1 Then
                ReDim lines(1 To layout.Count)
                For i = 1 To layout.Count
                    lines(i) = "PERF_BITS|" & CStr(n) & "|" & layout.BarID(i) & "|" & TestPerformanceBits(layout.X(i)) & "|" & _
                        TestPerformanceBits(layout.Y(i)) & "|" & TestPerformanceBits(layout.Diameter(i))
                Next i
                mReport = mReport & Join(lines, vbCrLf) & vbCrLf
            End If
        Next repeat
    Next n
    GoTo Finished
Failed:
    TestCheck "performance.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunIndividualContourPerformanceTests = mReport & "TOTAL_INDIVIDUAL_CONTOUR_PERFORMANCE: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает исходные байты Double, не преобразуя число в округлённую строку.
Private Function TestPerformanceBits(ByVal value As Double) As String
    Dim number As TPerformanceReal, bytes As TPerformanceBytes, i As Long
    number.Value = value: LSet bytes = number
    For i = 0 To 7: TestPerformanceBits = TestPerformanceBits & Right$("0" & Hex$(bytes.Value(i)), 2): Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' объединение соседних рядов сохраняет первого владельца и только фактические подписи.
Public Function RunOffsetNeighborMergeTests() As String
    Dim kind As Long, reverse As Long, withMain As Long
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    For kind = 0 To 3
        For reverse = 0 To 1
            For withMain = 0 To 1
                TestOffsetNeighborMerge kind, reverse = 1, withMain = 1
            Next withMain
        Next reverse
    Next kind
    TestNonAdjacentRows
    GoTo Finished
Failed:
    TestCheck "neighbor.merge.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunOffsetNeighborMergeTests = mReport & "TOTAL_OFFSET_NEIGHBOR_MERGE: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' пример пользователя 300 x 500, R100, полные отступы сверху 60 и сбоку 20 мм.
Public Function RunUserRoundedContourTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    TestUserRoundedContour False
    TestUserRoundedContour True
    TestCollapsedUserCorners 100#, False
    TestCollapsedUserCorners 120#, False
    TestCollapsedUserCorners 100#, True
    TestCollapsedUserCorners 120#, True
    GoTo Finished
Failed:
    TestCheck "user.rounded.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunUserRoundedContourTests = mReport & "TOTAL_USER_ROUNDED_CONTOUR: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' недостающие метаданные, неоднозначный as и невозможный замкнутый контур.
Public Function RunIndividualContourFailureTests() As String
    Dim kind As Long, facade As CSectionDefinitionBuilder, boundary As CGeometryPath, path As CGeometryPath
    Dim options As CRebarRuleOptions, i As Long, id As String, definition As CSectionDefinition, link As TRebarContourLink
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    For kind = 0 To 2
        Set facade = New CSectionDefinitionBuilder: facade.Initialize "IndividualFailure"
        Set boundary = facade.RectanglePath("Boundary", 0#, 0#, 1000#, 800#): facade.Outer "Outer", boundary.ID
        For i = 1 To IIf(kind = 2, 4, 1)
            Select Case i
                Case 1: Set path = facade.LinePath("Top", 100#, IIf(kind = 1, 400#, 750#), 900#, IIf(kind = 1, 400#, 750#))
                Case 2: Set path = facade.LinePath("Right", 950#, 700#, 950#, 100#)
                Case 3: Set path = facade.LinePath("Bottom", 900#, 50#, 100#, 50#)
                Case 4: Set path = facade.LinePath("Left", 50#, 100#, 50#, 700#)
            End Select
            id = "Side" & CStr(i)
            Set options = New CRebarRuleOptions: options.SourceReference = "IndividualFailure[Main," & id & "]"
            facade.RebarByCount id & ".Main", path.ID, 3, 10#, options
            If kind > 0 Then facade.RebarContour id & ".Main", "Outer", boundary.ID, IIf(kind = 1, 400#, 50#), id & ".Second", id & ".Third"
            TestAddRows facade, id, -1#, IIf(kind = 2, 600#, 40#), IIf(kind = 2, 650#, 70#), False
        Next i
        Set definition = facade.Finish
        If definition.RebarContourCount > 0 Then
            definition.CopyRebarContour 1, link
            link.MainRuleID = "ChangedCopy"
            definition.CopyRebarContour 1, link
            TestCheck "individual.metadata.copy.immutable", link.MainRuleID = "Side1.Main"
        End If
        TestExpectContourFailure definition, Choose(kind + 1, "RebarContour", "однозначно", "контур")
    Next kind
    GoTo Finished
Failed:
    TestCheck "individual.failure.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunIndividualContourFailureTests = mReport & "TOTAL_INDIVIDUAL_CONTOUR_FAILURES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' реальные рецепты OneRect, TwoRect, PointedSection, HollowRectangle и Circle.
Public Function RunIndividualContourWorkbookTests() As String
    Dim config As Object, saved As Variant, fixtures As Variant, fixture As String, table As Object, tableName As String
    Dim keyed As CSectionInputReader, keys As Variant, key As Variant, side As Long, scenario As Long, active As Boolean
    Dim settings As CSystemSettingsReader, units As CUnitSystem, reader As CSectionInputSnapshotReader, registry As CSectionTypeRegistry
    Dim context As CSectionRecipeContext, definition As CSectionDefinition, executor As CRebarRuleExecutor, layout As CRebarLayout
    Dim source As TRebarPosition, rule As TRebarRule, i As Long, owner As Long, expectedX As Double, expectedY As Double, top2 As Long, top3 As Long
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    ThisWorkbook.Names("rngUnitSettings").RefersToRange.Cells(2, 2).Value2 = "mm"
    fixtures = Array("Rectangle", "Rectangle", "LSection", "RoundedMixed", "HollowOffset", "HollowOffset", "Circle", "HollowOffset")
    For scenario = 0 To UBound(fixtures)
        fixture = CStr(fixtures(scenario)): modTestShapeArchitecture.TestConfigureShapeFixture fixture
        Select Case scenario
            Case 0, 1, 2
                tableName = "rngRectanglesGeometry"
                Set table = ThisWorkbook.Names(tableName).RefersToRange
                TestFixtureInputCell(table, "General", "Shape").Value2 = IIf(scenario = 2, "TwoRect", "OneRect")
                TestFixtureInputCell(table, "General", "B1").Value2 = IIf(scenario = 2, 1400#, 2400#)
                TestFixtureInputCell(table, "General", "H1").Value2 = IIf(scenario = 2, 900#, 1800#)
                TestFixtureInputCell(table, "General", "B2").Value2 = 2400#: TestFixtureInputCell(table, "General", "H2").Value2 = 800#
                TestFixtureInputCell(table, "General", "XOffset").Value2 = 400#
            Case 3
                tableName = "rngPointedSectionGeometry": Set table = ThisWorkbook.Names(tableName).RefersToRange
                TestFixtureInputCell(table, "General", "B").Value2 = 2000#: TestFixtureInputCell(table, "General", "H").Value2 = 2400#
                For Each key In Array("Left.", "Right.")
                    TestFixtureInputCell(table, "General", CStr(key) & "W").Value2 = 400#
                    TestFixtureInputCell(table, "General", CStr(key) & "R1").Value2 = 400#
                    TestFixtureInputCell(table, "General", CStr(key) & "R2").Value2 = 400#
                Next key
            Case 6
                tableName = "rngCircleGeometry": Set table = ThisWorkbook.Names(tableName).RefersToRange
                TestFixtureInputCell(table, "General", "D").Value2 = 2400#
            Case Else
                tableName = "rngHollowRectangleGeometry": Set table = ThisWorkbook.Names(tableName).RefersToRange
                TestFixtureInputCell(table, "General", "B").Value2 = 2400#: TestFixtureInputCell(table, "General", "H").Value2 = 2400#
                TestFixtureInputCell(table, "General", "R").Value2 = 300#
                TestFixtureInputCell(table, "General", "Bo").Value2 = 800#: TestFixtureInputCell(table, "General", "Ho").Value2 = 800#
                TestFixtureInputCell(table, "General", "Ro").Value2 = IIf(scenario = 4, 150#, 0#)
                TestFixtureInputCell(table, "General", "Xo").Value2 = 0#: TestFixtureInputCell(table, "General", "Yo").Value2 = 0#
        End Select
        Set keyed = TestFixtureInput(table.Value2): keys = keyed.RowKeys("Main"): side = 0
        For Each key In keys
            side = side + 1
            TestFixtureInputCell(table, "Main", CStr(key), "as").Value2 = IIf(scenario <= 1, 50#, 50# + 5# * side)
            TestFixtureInputCell(table, "Main", CStr(key), "d").Value2 = 20#
            If Left$(CStr(key), 7) <> "Opening" Then TestFixtureInputCell(table, "Main", CStr(key), "n").Value2 = IIf(scenario = 6, 12, 5)
            TestFixtureInputCell(table, "Main", CStr(key), "s0").Value2 = IIf(scenario = 6, 0#, 150#)
            TestFixtureInputCell(table, "Main", CStr(key), "s1").Value2 = IIf(scenario = 6, 0#, 150#)
            active = (scenario <> 0 Or CStr(key) = "R1.Top")
            If scenario = 7 And (CStr(key) = "OpeningLeft" Or CStr(key) = "OpeningRight") Then
                TestFixtureInputCell(table, "Main", CStr(key), "d").Value2 = 0#: active = False
            End If
            TestFixtureInputCell(table, "Extra", CStr(key), "d2").Value2 = IIf(active, 10#, 0#)
            TestFixtureInputCell(table, "Extra", CStr(key), "d3").Value2 = IIf(active, 10#, 0#)
            TestFixtureInputCell(table, "Extra", CStr(key), "a2").Value2 = IIf(scenario <= 1, 150#, 100#)
            TestFixtureInputCell(table, "Extra", CStr(key), "a3").Value2 = 200#
            TestFixtureInputCell(table, "Extra", CStr(key), "pos2").Value2 = "Stacked"
            TestFixtureInputCell(table, "Extra", CStr(key), "pos3").Value2 = "Stacked"
            TestFixtureInputCell(table, "Extra", CStr(key), "bind2").Value2 = "EachBar"
            TestFixtureInputCell(table, "Extra", CStr(key), "bind3").Value2 = "EachBar"
        Next key
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set registry = New CSectionTypeRegistry: Set reader = New CSectionInputSnapshotReader
        fixture = settings.GetRequiredString("Geometry.Type")
        Set context = reader.Capture(ThisWorkbook, registry.RecipeTableNames(fixture), units)
        Set definition = registry.CreateDefinition(fixture, context)
        Set executor = New CRebarRuleExecutor: Set layout = TestBuildLayout(definition, executor)
        TestCheck "recipe." & CStr(scenario) & ".created", layout.Count > 0
        top2 = 0: top3 = 0
        For i = 1 To executor.PositionCount
            definition.CopyRule definition.RuleIndex(executor.PositionRuleID(i)), rule
            If rule.Mode = DependentOffset Then
                executor.TestCopyTransfer i, source, rule: owner = executor.PositionPhysicalOwner(i)
                expectedX = source.X + rule.TangentOffset * source.TangentX - rule.NormalOffset * source.TangentY
                expectedY = source.Y + rule.TangentOffset * source.TangentY + rule.NormalOffset * source.TangentX
                TestNear "recipe." & CStr(scenario) & ".preserve.x", layout.X(owner), expectedX
                TestNear "recipe." & CStr(scenario) & ".preserve.y", layout.Y(owner), expectedY
                If rule.RuleID = "B1.as_1.Row2" Then top2 = top2 + 1
                If rule.RuleID = "B1.as_1.Row3" Then top3 = top3 + 1
            End If
        Next i
        If scenario <= 1 Then
            TestCheck "recipe.OneRect.top2", top2 = IIf(scenario = 0, 5, 3)
            TestCheck "recipe.OneRect.top3", top3 = IIf(scenario = 0, 5, 3)
        End If
        mReport = mReport & "RECIPE_CASE: " & CStr(scenario) & "; " & fixture & "; bars=" & CStr(layout.Count) & vbCrLf & executor.DiagnosticsText
    Next scenario
    GoTo Finished
Failed:
    TestCheck "individual.workbook.runtime.case" & CStr(scenario) & "; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    RunIndividualContourWorkbookTests = mReport & "TOTAL_INDIVIDUAL_CONTOUR_WORKBOOK: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет причину, фактическую привязку полей и отсутствие маскировки ошибки пропусками.
Private Sub TestExpectContourFailure(ByVal definition As CSectionDefinition, ByVal expected As String)
    Dim executor As CRebarRuleExecutor, layout As CRebarLayout, reason As String
    Set executor = New CRebarRuleExecutor
    On Error GoTo Rejected
    Set layout = TestBuildLayout(definition, executor)
    TestCheck "individual.failure.must.reject", False
    Exit Sub
Rejected:
    reason = Err.Description
    TestCheck "individual.failure.reason." & expected, InStr(1, reason, expected, vbTextCompare) > 0
    TestCheck "individual.failure.source", InStr(reason, "IndividualRectangle[Extra,") > 0 Or InStr(reason, "IndividualFailure[Main,") > 0
    TestCheck "individual.failure.not.mass.skip", Not TestHasOffsetSkip(executor)
    mReport = mReport & "REJECTED: " & reason & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Верхние 150/200 мм не переносятся на соседние стороны без дополнительных рядов.
' При включении всех сторон крайние позиции проверяются по их собственным отступам.
Private Sub TestRectangle(ByVal allSides As Boolean, ByVal reverse As Boolean, ByVal everySecond As Boolean)
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, boundary As CGeometryPath
    Dim options As CRebarRuleOptions, definition As CSectionDefinition, executor As CRebarRuleExecutor, layout As CRebarLayout
    Dim i As Long, sign As Double, expected As Long, row2 As Long, row3 As Long, owner As Long, ruleID As String
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "IndividualRectangle"
    Set boundary = facade.RectanglePath("Boundary", 0#, 0#, 1000#, 800#)
    If reverse Then
        Set boundary = boundary.Reversed.WithID("Boundary.Reverse"): facade.AddPath boundary
    End If
    facade.Outer "Outer", boundary.ID
    For i = 1 To 4
        Select Case i
            Case 1: Set path = facade.LinePath("Top", 50#, 750#, 950#, 750#)
            Case 2: Set path = facade.LinePath("Right", 950#, 750#, 950#, 50#)
            Case 3: Set path = facade.LinePath("Bottom", 950#, 50#, 50#, 50#)
            Case 4: Set path = facade.LinePath("Left", 50#, 50#, 50#, 750#)
        End Select
        sign = -1#
        If reverse Then
            Set path = path.Reversed.WithID(path.ID & ".Reverse"): facade.AddPath path: sign = 1#
        End If
        ruleID = "Side" & CStr(i)
        Set options = New CRebarRuleOptions: options.Coincidence = KeepPublishedOwner
        options.SourceReference = "IndividualRectangle[Main," & CStr(i) & "]"
        facade.RebarByCount ruleID & ".Main", path.ID, 5, 10#, options
        facade.RebarContour ruleID & ".Main", "Outer", boundary.ID, 50#, ruleID & ".Second", ruleID & ".Third"
        If i = 1 Or allSides Then TestAddRows facade, ruleID, sign, 150#, 200#, everySecond
    Next i
    Set definition = facade.Finish: Set executor = New CRebarRuleExecutor
    Set layout = TestBuildLayout(definition, executor)
    For i = 1 To executor.PositionCount
        owner = executor.PositionPhysicalOwner(i)
        If executor.PositionRuleID(i) = "Side1.Second" Then
            row2 = row2 + 1: TestNear "rectangle.row2.y", layout.Y(owner), 600#
        ElseIf executor.PositionRuleID(i) = "Side1.Third" Then
            row3 = row3 + 1: TestNear "rectangle.row3.y", layout.Y(owner), 550#
        End If
    Next i
    expected = 5: If allSides Or everySecond Then expected = 3
    If allSides And everySecond Then expected = 1
    TestCheck "rectangle.top.row2.count", row2 = expected
    TestCheck "rectangle.top.row3.count", row3 = expected
    If allSides Then
        TestCheck "rectangle.excluded.ends.diagnostic", TestHasOffsetSkip(executor)
        TestCheck "rectangle.excluded.ends.address", InStr(executor.DiagnosticsText, "IndividualRectangle[Extra,") > 0
    Else
        TestCheck "rectangle.no.false.corner.skip", Not TestHasOffsetSkip(executor)
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Пять позиций верха: оба конца, середины обеих дуг и середина линии.
' Верхние дополнительные ряды имеют 40/70 мм, боковые 60/90 мм, as также различны.
Private Sub TestRoundedCorners(ByVal reverse As Boolean)
    Dim facade As CSectionDefinitionBuilder, polygon As CPolygonSpec, boundary As CGeometryPath, whole As CGeometryPath, path As CGeometryPath
    Dim query As CSectionGeometryQuery, options As CRebarRuleOptions, definition As CSectionDefinition
    Dim executor As CRebarRuleExecutor, layout As CRebarLayout, stations As Variant, edge As TRegionEdge
    Dim i As Long, sign As Double, distances As Variant, ruleID As String, row2 As Long, row3 As Long
    Dim x As Double, y As Double, tx As Double, ty As Double, normal As Double, owner As Long, source As Long, expectedX As Double, expectedY As Double
    Dim parent As TRebarPosition, transfer As TRebarRule, segment As String, parameter As Double, expectedSegment As String
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "IndividualRounded"
    Set polygon = New CPolygonSpec
    polygon.AddVertex "BL", 0#, 0#, 120#: polygon.AddVertex "BR", 1000#, 0#, 140#
    polygon.AddVertex "TR", 1000#, 800#, 180#: polygon.AddVertex "TL", 0#, 800#, 160#: polygon.Seal
    Set boundary = facade.RoundedPolygon("Boundary", polygon)
    If reverse Then
        Set boundary = boundary.Reversed.WithID("Boundary.Reverse"): facade.AddPath boundary
    End If
    facade.Outer "Outer", boundary.ID
    distances = Array(50#, 65#, 45#)
    For i = 1 To 3
        Set whole = facade.OffsetRoundedPolygon("Whole" & CStr(i), polygon, CDbl(distances(i - 1)))
        Select Case i
            Case 1
                Set path = facade.Chain("Top", whole.ID, "Edge.BR", 1#, "Edge.TL", 0#)
                path.CopySegment 1, edge: normal = modGeometryPrimitives.EdgeLength(edge) / 2#
                path.CopySegment path.SegmentCount, edge
                stations = Array(0#, normal, path.Length / 2#, path.Length - modGeometryPrimitives.EdgeLength(edge) / 2#, path.Length)
            Case 2: Set path = facade.Chain("Right", whole.ID, "Edge.BR", 0#, "Edge.BR", 1#)
            Case 3: Set path = facade.Chain("Left", whole.ID, "Edge.TL", 0#, "Edge.TL", 1#)
        End Select
        sign = 1#
        If reverse Then
            Set path = path.Reversed.WithID(path.ID & ".Reverse"): facade.AddPath path: sign = -1#
            If i = 1 Then stations = Array(0#, path.Length - CDbl(stations(3)), path.Length - CDbl(stations(2)), path.Length - CDbl(stations(1)), path.Length)
        End If
        ruleID = "Face" & CStr(i)
        Set options = New CRebarRuleOptions: options.SourceReference = "IndividualRounded[Main," & CStr(i) & "]"
        If i = 1 Then
            facade.RebarAtStations ruleID & ".Main", path.ID, stations, 10#, options
        Else
            options.EndMode = ExcludeBoth
            facade.RebarByCount ruleID & ".Main", path.ID, 3, 10#, options
        End If
        facade.RebarContour ruleID & ".Main", "Outer", boundary.ID, CDbl(distances(i - 1)), ruleID & ".Second", ruleID & ".Third"
        If i = 1 Then TestAddRows facade, ruleID, sign, 40#, 70#, False Else TestAddRows facade, ruleID, sign, 60#, 90#, False
    Next i
    Set definition = facade.Finish: Set executor = New CRebarRuleExecutor: Set layout = TestBuildLayout(definition, executor)
    For i = 1 To executor.PositionCount
        ruleID = executor.PositionRuleID(i)
        If ruleID = "Face1.Second" Or ruleID = "Face1.Third" Then
            If ruleID = "Face1.Second" Then normal = 40#: row2 = row2 + 1 Else normal = 70#: row3 = row3 + 1
            Set path = definition.Path(IIf(reverse, "Top.Reverse", "Top"))
            If ruleID = "Face1.Second" Then source = row2 Else source = row3
            path.AtDistance CDbl(stations(source - 1)), x, y, tx, ty
            If reverse Then normal = -normal
            expectedX = x - normal * ty: expectedY = y + normal * tx
            owner = executor.PositionPhysicalOwner(i)
            TestNear "rounded.retained.x", layout.X(owner), expectedX: TestNear "rounded.retained.y", layout.Y(owner), expectedY
            executor.TestCopyTransfer i, parent, transfer
            TestCheck "rounded.owner.Main", parent.RuleIndex = definition.RuleIndex("Face1.Main")
            path.ResolveAnchor parent.X, parent.Y, segment, parameter
            expectedSegment = "Edge.TR"
            If source < 3 Then expectedSegment = IIf(reverse, "Corner.TL", "Corner.TR")
            If source > 3 Then expectedSegment = IIf(reverse, "Corner.TR", "Corner.TL")
            TestCheck "rounded.owner.segment", path.SourceSegmentID(path.SegmentIndex(segment)) = expectedSegment
            mReport = mReport & "CORNER: " & CStr(reverse) & "; " & ruleID & "; source=" & CStr(source) & "; segment=" & expectedSegment & "; x=" & CStr(layout.X(owner)) & "; y=" & CStr(layout.Y(owner)) & vbCrLf
        End If
    Next i
    TestCheck "rounded.row2.all.five", row2 = 5: TestCheck "rounded.row3.all.five", row3 = 5
    TestCheck "rounded.no.artificial.corner.skip", Not TestHasOffsetSkip(executor)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' При совпадении Main с серединой сечения исходную сторону однозначно задаёт переданная нормаль.
Private Sub TestMidplaneMetadata(ByVal sign As Double)
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim definition As CSectionDefinition, executor As CRebarRuleExecutor, layout As CRebarLayout, i As Long, count As Long
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "MidplaneMetadata"
    Set path = facade.RectanglePath("Boundary", 0#, 0#, 1000#, 800#): facade.Outer "Outer", path.ID
    Set path = facade.LinePath("Middle", 100#, 400#, 900#, 400#)
    Set options = New CRebarRuleOptions: options.SourceReference = "MidplaneMetadata[Main]"
    facade.RebarByCount "Main", path.ID, 3, 10#, options
    facade.RebarContour "Main", "Outer", "Boundary", 400#, "Second", "Third", normalSign:=sign
    Set options = New CRebarRuleOptions: options.NormalOffset = sign * 40#: options.ValidateRowOffset = True
    options.SourceReference = "MidplaneMetadata[Extra,2]"
    facade.RebarDependent "Second", "Main", 10#, options
    Set definition = facade.Finish: Set executor = New CRebarRuleExecutor: Set layout = TestBuildLayout(definition, executor)
    For i = 1 To executor.PositionCount
        If executor.PositionRuleID(i) = "Second" Then
            count = count + 1: TestNear "midplane.explicit.normal.y", layout.Y(executor.PositionPhysicalOwner(i)), 400# + sign * 40#
        End If
    Next i
    TestCheck "midplane.explicit.normal.count", count = 3
    TestCheck "midplane.no.false.skip", Not TestHasOffsetSkip(executor)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отдельные области двух наружных границ, двух отверстий и обоих дополнительных рядов.
' У первого отверстия верх смещён дополнительно, у второго остаётся только известный as.
Private Sub TestIndependentContourDomains()
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim definition As CSectionDefinition, contours As CRebarOffsetContours, rules() As TRebarRule
    Dim i As Long, row As Long, loopID As String, mainID As String, offsetX As Double
    Dim a2 As Long, a3 As Long, b2 As Long, b3 As Long, hole2 As Long, hole3 As Long
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "IndependentContourDomains"
    For i = 1 To 2
        loopID = "Outer" & CStr(i): offsetX = CDbl(i - 1) * 2000#
        Set path = facade.RectanglePath(loopID & ".Boundary", offsetX, 0#, 1000#, 800#)
        facade.Outer loopID, path.ID
    Next i
    For i = 1 To 2
        loopID = "Opening" & CStr(i)
        Set path = facade.RectanglePath(loopID & ".Boundary", 150# + CDbl(i - 1) * 500#, 250#, 200#, 200#, "Bottom", "Right", "Top", "Left")
        facade.Opening loopID, path.ID, "Outer1"
        facade.RebarContourBaseline loopID, path.ID, Array("Bottom", "Right", "Top", "Left"), 20#, "DomainScope[Opening.as]"
    Next i
    For i = 1 To 2
        loopID = "Outer" & CStr(i): mainID = loopID & ".Main": offsetX = CDbl(i - 1) * 2000#
        Set path = facade.LinePath(loopID & ".Axis", offsetX + 100#, 750#, offsetX + 900#, 750#)
        Set options = New CRebarRuleOptions: options.SourceReference = "DomainScope[" & mainID & "]"
        facade.RebarByCount mainID, path.ID, 3, 10#, options
        facade.RebarContour mainID, loopID, loopID & ".Boundary", 50#, loopID & ".R2", loopID & ".R3", normalSign:=-1#
        For row = 2 To 3
            Set options = New CRebarRuleOptions: options.ValidateRowOffset = True
            options.NormalOffset = -CDbl(IIf(i = 1, 150#, 50#)) - CDbl(row - 2) * 50#
            options.SourceReference = "DomainScope[" & loopID & ".R" & CStr(row) & "]"
            facade.RebarDependent loopID & ".R" & CStr(row), mainID, 8#, options
        Next row
    Next i
    Set path = facade.LinePath("Opening1.Axis", 175#, 470#, 325#, 470#)
    Set options = New CRebarRuleOptions: options.SourceReference = "DomainScope[Opening1.Main]"
    facade.RebarByCount "Opening1.Main", path.ID, 3, 10#, options
    facade.RebarContour "Opening1.Main", "Opening1", "Opening1.Boundary", 20#, "Opening1.R2", "Opening1.R3", Array("Top"), 1#
    For row = 2 To 3
        Set options = New CRebarRuleOptions: options.ValidateRowOffset = True: options.NormalOffset = CDbl(row - 1) * 30#
        options.SourceReference = "DomainScope[Opening1.R" & CStr(row) & "]"
        facade.RebarDependent "Opening1.R" & CStr(row), "Opening1.Main", 8#, options
    Next row
    Set definition = facade.Finish
    ReDim rules(1 To definition.RuleCount)
    For i = 1 To definition.RuleCount: definition.CopyRule i, rules(i): Next i
    Set contours = New CRebarOffsetContours: contours.Initialize definition, rules
    a2 = definition.RuleIndex("Outer1.R2"): a3 = definition.RuleIndex("Outer1.R3")
    b2 = definition.RuleIndex("Outer2.R2"): b3 = definition.RuleIndex("Outer2.R3")
    hole2 = definition.RuleIndex("Opening1.R2"): hole3 = definition.RuleIndex("Opening1.R3")
    TestCheck "domain.first.outer.boundary", contours.Accepts(a2, 500#, 600#)
    TestCheck "domain.first.outer.above", Not contours.Accepts(a2, 500#, 650#)
    TestCheck "domain.second.outer.inside", contours.Accepts(b2, 2500#, 650#)
    TestCheck "domain.second.outer.above", Not contours.Accepts(b2, 2500#, 750#)
    TestCheck "domain.first.not.second", Not contours.Accepts(a2, 2500#, 650#)
    TestCheck "domain.second.not.first", Not contours.Accepts(b2, 500#, 550#)
    TestCheck "domain.first.opening.center", Not contours.Accepts(a2, 250#, 350#)
    TestCheck "domain.second.opening.center", Not contours.Accepts(a2, 750#, 350#)
    TestCheck "domain.first.opening.extra.offset", Not contours.Accepts(a2, 250#, 480#)
    TestCheck "domain.second.opening.only.as", contours.Accepts(a2, 750#, 480#)
    TestCheck "domain.opening.shares.outer.limit", Not contours.Accepts(hole2, 500#, 650#)
    TestCheck "domain.opening.own.boundary", contours.Accepts(hole2, 250#, 500#)
    TestCheck "domain.opening.sees.other.hole", Not contours.Accepts(hole2, 750#, 350#)
    TestCheck "domain.first.rows.independent", contours.Accepts(a2, 500#, 575#) And Not contours.Accepts(a3, 500#, 575#)
    TestCheck "domain.second.rows.independent", contours.Accepts(b2, 2500#, 675#) And Not contours.Accepts(b3, 2500#, 675#)
    TestCheck "domain.opening.third.boundary", contours.Accepts(hole3, 250#, 530#)
    TestCheck "domain.opening.rows.independent", contours.Accepts(hole2, 250#, 515#) And Not contours.Accepts(hole3, 250#, 515#)
    mReport = mReport & "DOMAIN_SCOPE: outers=2; openings=2; rows=2,3; independent=True" & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет точные дуги и прямые перемычки, затем сохранение пяти стержней верхнего ряда.
Private Sub TestUserRoundedContour(ByVal reverse As Boolean)
    Dim facade As CSectionDefinitionBuilder, polygon As CPolygonSpec, boundary As CGeometryPath, whole As CGeometryPath, path As CGeometryPath
    Dim edges() As TRegionEdge, distances() As Double, i As Long, count As Long, connectorCount As Long, arcCount As Long
    Dim offset As CGeometryVariableOffset, region As CConcreteRegion, query As CSectionGeometryQuery, edge As TRegionEdge
    Dim options As CRebarRuleOptions, definition As CSectionDefinition, executor As CRebarRuleExecutor, layout As CRebarLayout
    Dim stations As Variant, halfArc As Double, rowCount As Long, sign As Double, parent As TRebarPosition, rule As TRebarRule, owner As Long
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "UserRounded300x500"
    Set polygon = New CPolygonSpec
    polygon.AddVertex "BL", 0#, 0#, 100#: polygon.AddVertex "BR", 300#, 0#, 100#
    polygon.AddVertex "TR", 300#, 500#, 100#: polygon.AddVertex "TL", 0#, 500#, 100#: polygon.Seal
    Set boundary = facade.RoundedPolygon("Boundary", polygon)
    count = boundary.SegmentCount: ReDim edges(1 To count): ReDim distances(1 To count)
    For i = 1 To count
        boundary.CopySegment i, edges(i): distances(i) = 20#
        Select Case boundary.SegmentID(i)
            Case "Corner.TR", "Edge.TR", "Corner.TL": distances(i) = 60#
        End Select
    Next i
    Set offset = New CGeometryVariableOffset: Set query = New CSectionGeometryQuery
    Set region = offset.Build(edges, distances, count, "UserRounded300x500[Top=60,Sides=20]")
    TestCheck "user.rounded.closed.one.loop", region.LoopCount = 1
    region.CopyEdges edges, count
    For i = 1 To count
        edge = edges(i)
        If edge.Sweep <> 0# Then
            arcCount = arcCount + 1
            If Abs(edge.CenterY - 400#) < GEOM_BOUNDARY_TOLERANCE Then _
                TestNear "user.rounded.upper.radius", Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2), 40#
        ElseIf Abs(edge.Y1 - 400#) < GEOM_BOUNDARY_TOLERANCE And Abs(edge.Y2 - 400#) < GEOM_BOUNDARY_TOLERANCE Then
            connectorCount = connectorCount + 1: TestNear "user.rounded.connector.length", Abs(edge.X2 - edge.X1), 40#
            mReport = mReport & "USER_CONNECTOR: (" & CStr(edge.X1) & "," & CStr(edge.Y1) & ") -> (" & CStr(edge.X2) & "," & CStr(edge.Y2) & ")" & vbCrLf
        End If
    Next i
    TestCheck "user.rounded.four.exact.arcs", arcCount = 4
    TestCheck "user.rounded.two.connectors", connectorCount = 2
    TestCheck "user.rounded.corner.inside", query.ContainsPoint(region, 200# + 40# / Sqr(2#), 400# + 40# / Sqr(2#), True)
    TestCheck "user.rounded.step.inside", query.ContainsPoint(region, 260#, 390#, True)
    TestCheck "user.rounded.step.outside", Not query.ContainsPoint(region, 260#, 410#, True)
    If reverse Then
        Set boundary = boundary.Reversed.WithID("Boundary.Reverse"): facade.AddPath boundary
    End If
    facade.Outer "Outer", boundary.ID
    Set whole = facade.OffsetRoundedPolygon("MainContour", polygon, 20#)
    For i = 1 To 4
        Select Case i
            Case 1
                Set path = facade.Chain("Top", whole.ID, "Edge.BR", 1#, "Edge.TL", 0#)
                path.CopySegment 1, edge: halfArc = modGeometryPrimitives.EdgeLength(edge) / 2#
                stations = Array(0#, halfArc, path.Length / 2#, path.Length - halfArc, path.Length)
            Case 2: Set path = facade.Chain("Right", whole.ID, "Edge.BR", 0#, "Edge.BR", 1#)
            Case 3: Set path = facade.Chain("Bottom", whole.ID, "Edge.BL", 0#, "Edge.BL", 1#)
            Case 4: Set path = facade.Chain("Left", whole.ID, "Edge.TL", 0#, "Edge.TL", 1#)
        End Select
        sign = 1#
        If reverse Then
            Set path = path.Reversed.WithID(path.ID & ".Reverse"): facade.AddPath path: sign = -1#
        End If
        Set options = New CRebarRuleOptions: options.SourceReference = "UserRounded300x500[Main," & CStr(i) & "]"
        If i = 1 Then
            facade.RebarAtStations "Top.Main", path.ID, stations, 4#, options
            facade.RebarContour "Top.Main", "Outer", boundary.ID, 20#, "Top.Second", "", normalSign:=sign
            Set options = New CRebarRuleOptions: options.NormalOffset = sign * 40#: options.ValidateRowOffset = True
            options.SourceReference = "UserRounded300x500[Extra,Top,2]"
            facade.RebarDependent "Top.Second", "Top.Main", 4#, options
        Else
            options.EndMode = ExcludeBoth
            facade.RebarByCount "Side" & CStr(i), path.ID, 3, 4#, options
            facade.RebarContour "Side" & CStr(i), "Outer", boundary.ID, 20#, "", "", normalSign:=sign
        End If
    Next i
    Set definition = facade.Finish: Set executor = New CRebarRuleExecutor: Set layout = TestBuildLayout(definition, executor)
    For i = 1 To executor.PositionCount
        If executor.PositionRuleID(i) = "Top.Second" Then
            rowCount = rowCount + 1: owner = executor.PositionPhysicalOwner(i)
            executor.TestCopyTransfer i, parent, rule
            TestNear "user.rounded.preserved.x", layout.X(owner), parent.X - rule.NormalOffset * parent.TangentY
            TestNear "user.rounded.preserved.y", layout.Y(owner), parent.Y + rule.NormalOffset * parent.TangentX
        End If
    Next i
    TestCheck "user.rounded.all.five.retained", rowCount = 5
    TestCheck "user.rounded.no.false.skip", Not TestHasOffsetSkip(executor)
    mReport = mReport & "USER_CASE: reverse=" & CStr(reverse) & "; top.R2=" & CStr(rowCount) & "; R=40; connectors=40" & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Совпадение, частичное пересечение, касание и разные номера R2/R3 у соседних граней.
' Подписи формируются штатным построителем из физического CRebarLayout.
Private Sub TestOffsetNeighborMerge(ByVal kind As Long, ByVal rightFirst As Boolean, ByVal withMain As Boolean)
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions, order As Variant
    Dim item As Variant, id As String, ruleID As String, x As Double, y As Double, diameter As Double, expectedDiameter As Double
    Dim definition As CSectionDefinition, executor As CRebarRuleExecutor, layout As CRebarLayout, i As Long, owner As Long
    Dim labels As CRebarGroupAnnotationBuilder, annotations As CSectionAnnotations, anchors(1 To 2) As TSectionAnnotation
    Dim expectedOwner As String, textTop As String, textRight As String, diagnostic As CGeometryDiagnostic, warnings As Long
    Dim expectedTop As String, expectedRight As String, totalArea As Double, expectedArea As Double
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "NeighborMerge"
    Set path = facade.RectanglePath("Boundary", 0#, 0#, 1000#, 800#): facade.Outer "Outer", path.ID
    order = Array("Top", "Right"): If rightFirst Then order = Array("Right", "Top")
    For Each item In order
        id = CStr(item)
        Set options = New CRebarRuleOptions: options.SourceReference = "NeighborMerge[Main," & id & "]"
        If withMain Then options.AnnotationGroup = id
        If id = "Top" Then
            x = 800#: If kind = 1 Or kind = 3 Then x = 795#
            If kind = 2 Then x = 782#
            Set path = facade.LinePath("TopLine", 50#, 750#, 950#, 750#)
            facade.RebarAtStations id & ".Main", path.ID, Array(x - 50#), 10#, options
            ruleID = id & ".Second": diameter = 16#
        Else
            y = 600#: If kind = 1 Or kind = 3 Then y = 595#
            Set path = facade.LinePath("RightLine", 950#, 750#, 950#, 50#)
            facade.RebarAtStations id & ".Main", path.ID, Array(750# - y), 10#, options
            ruleID = id & ".Second": If kind = 3 Then ruleID = id & ".Third"
            diameter = 20#
        End If
        facade.RebarContour id & ".Main", "Outer", "Boundary", 50#, id & ".Second", id & ".Third", normalSign:=-1#
        Set options = New CRebarRuleOptions: options.NormalOffset = -150#: options.ValidateRowOffset = True
        options.AnnotationGroup = id: options.SourceReference = "NeighborMerge[Extra," & id & "]"
        facade.RebarDependent ruleID, id & ".Main", diameter, options
    Next item
    Set definition = facade.Finish: Set executor = New CRebarRuleExecutor: Set layout = TestBuildLayout(definition, executor)
    TestCheck "neighbor.actual.bar.count", layout.Count = IIf(kind = 2, 4, 3)
    TestCheck "neighbor.merged.counter", executor.MergedPositions = IIf(kind = 2, 0, 1)
    TestCheck "neighbor.actual.overlaps", executor.BarOverlaps = 0
    expectedOwner = IIf(rightFirst, "Right", "Top"): expectedDiameter = IIf(rightFirst, 20#, 16#)
    If kind <> 2 Then
        For i = 1 To executor.PositionCount
            If InStr(executor.PositionRuleID(i), ".Main") = 0 Then
                If owner = 0 Then owner = executor.PositionPhysicalOwner(i)
                TestCheck "neighbor.same.physical.owner", executor.PositionPhysicalOwner(i) = owner
            End If
        Next i
        TestNear "neighbor.first.diameter", layout.Diameter(owner), expectedDiameter
        TestCheck "neighbor.first.annotation.group", layout.BarAnnotationGroupName(owner) = expectedOwner
        If rightFirst Then
            TestNear "neighbor.first.x", layout.X(owner), 800#: TestNear "neighbor.first.y", layout.Y(owner), y
        Else
            TestNear "neighbor.first.x", layout.X(owner), x: TestNear "neighbor.first.y", layout.Y(owner), 600#
        End If
        TestCheck "neighbor.summary.warning", InStr(executor.DiagnosticsText, "Предупреждение: объединено") > 0
        For i = 1 To executor.DiagnosticCount
            Set diagnostic = executor.DiagnosticAt(i)
            If diagnostic.Code = "OFFSET_NEIGHBOR_MERGED" Then
                warnings = warnings + 1
                TestCheck "neighbor.warning.both.rows", InStr(diagnostic.Reason, "Top.Second") > 0 And InStr(diagnostic.Reason, "Right.") > 0
                TestCheck "neighbor.warning.source", InStr(diagnostic.SourceReference, "NeighborMerge[Extra,") > 0
            End If
        Next i
        TestCheck "neighbor.one.warning", warnings = 1
    End If
    anchors(1).Name = "Top": anchors(1).EndX = 100#: anchors(1).NormalY = 1#
    anchors(2).Name = "Right": anchors(2).EndY = 100#: anchors(2).NormalX = 1#
    Set labels = New CRebarGroupAnnotationBuilder: Set annotations = New CSectionAnnotations
    labels.AddRebarLabels annotations, layout, anchors, 2
    TestCheck "neighbor.label.count", annotations.Count = IIf(kind = 2 Or withMain, 2, 1)
    For i = 1 To annotations.Count
        If annotations.AnnotationID(i) = "REBAR_Top" Then textTop = annotations.Text(i)
        If annotations.AnnotationID(i) = "REBAR_Right" Then textRight = annotations.Text(i)
    Next i
    If withMain Then expectedTop = "1" & ChrW$(216) & "10": expectedRight = expectedTop
    If kind = 2 Or Not rightFirst Then
        If Len(expectedTop) > 0 Then expectedTop = expectedTop & " + "
        expectedTop = expectedTop & "1" & ChrW$(216) & "16"
    End If
    If kind = 2 Or rightFirst Then
        If Len(expectedRight) > 0 Then expectedRight = expectedRight & " + "
        expectedRight = expectedRight & "1" & ChrW$(216) & "20"
    End If
    TestCheck "neighbor.label.top.exact.physical.count", textTop = expectedTop
    TestCheck "neighbor.label.right.exact.physical.count", textRight = expectedRight
    For i = 1 To layout.Count: totalArea = totalArea + layout.Area(i): Next i
    expectedArea = GEOM_PI * (2# * 10# ^ 2 + expectedDiameter ^ 2) / 4#
    If kind = 2 Then expectedArea = GEOM_PI * (2# * 10# ^ 2 + 16# ^ 2 + 20# ^ 2) / 4#
    TestNear "neighbor.area.only.physical.bars", totalArea, expectedArea
    If kind <> 2 Then
        TestCheck "neighbor.label.only.kept.diameter", (InStr(textTop & textRight, ChrW$(216) & CStr(expectedDiameter)) > 0)
        TestCheck "neighbor.label.no.discarded.diameter", InStr(textTop & textRight, ChrW$(216) & CStr(IIf(rightFirst, 16, 20))) = 0
        If Not withMain Then TestCheck "neighbor.empty.face.no.label", Len(IIf(rightFirst, textTop, textRight)) = 0
    End If
    mReport = mReport & "NEIGHBOR_MERGE: kind=" & CStr(kind) & "; rightFirst=" & CStr(rightFirst) & "; mainLabels=" & CStr(withMain) & _
        "; physical=" & CStr(layout.Count) & "; top=" & textTop & "; right=" & textRight & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Противоположные грани не являются соседними: их пересекающиеся ряды отклоняются.
Private Sub TestNonAdjacentRows()
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim definition As CSectionDefinition, executor As CRebarRuleExecutor, layout As CRebarLayout, i As Long, id As String, reason As String
    On Error GoTo Rejected
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "NonAdjacentRows"
    Set path = facade.RectanglePath("Boundary", 0#, 0#, 1000#, 800#): facade.Outer "Outer", path.ID
    For i = 1 To 2
        id = CStr(i)
        If i = 1 Then Set path = facade.LinePath(id, 50#, 750#, 950#, 750#) Else Set path = facade.LinePath(id, 950#, 50#, 50#, 50#)
        Set options = New CRebarRuleOptions: options.SourceReference = "NonAdjacentRows[Main," & id & "]"
        facade.RebarAtStations id & ".Main", path.ID, Array(450#), 10#, options
        facade.RebarContour id & ".Main", "Outer", "Boundary", 50#, id & ".Second", "", normalSign:=-1#
        Set options = New CRebarRuleOptions: options.ValidateRowOffset = True
        options.NormalOffset = -IIf(i = 1, 300#, 390#): options.SourceReference = "NonAdjacentRows[Extra," & id & "]"
        facade.RebarDependent id & ".Second", id & ".Main", IIf(i = 1, 16#, 20#), options
    Next i
    Set definition = facade.Finish: Set executor = New CRebarRuleExecutor: Set layout = TestBuildLayout(definition, executor)
    TestCheck "neighbor.nonadjacent.must.reject", False
    Exit Sub
Rejected:
    reason = Err.Description
    TestCheck "neighbor.nonadjacent.physical.overlap.rejected", InStr(reason, "Поперечные сечения стержней пересекаются") > 0
    TestCheck "neighbor.nonadjacent.no.merge", executor.MergedPositions = 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' При отступе 100/120 мм верхние R100 исчезают, а прямые стыкуются в острых углах.
Private Sub TestCollapsedUserCorners(ByVal topOffset As Double, ByVal reverse As Boolean)
    Dim facade As CSectionDefinitionBuilder, polygon As CPolygonSpec, path As CGeometryPath, query As CSectionGeometryQuery
    Dim offset As CGeometryVariableOffset, region As CConcreteRegion, edges() As TRegionEdge, distances() As Double
    Dim i As Long, count As Long, arcs As Long, upper As Long, direction As Double, segmentID As String
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "CollapsedUserCorners"
    Set polygon = New CPolygonSpec
    polygon.AddVertex "BL", 0#, 0#, 100#: polygon.AddVertex "BR", 300#, 0#, 100#
    polygon.AddVertex "TR", 300#, 500#, 100#: polygon.AddVertex "TL", 0#, 500#, 100#: polygon.Seal
    Set path = facade.RoundedPolygon("Boundary", polygon): direction = 1#
    If reverse Then Set path = path.Reversed: direction = -1#
    count = path.SegmentCount: ReDim edges(1 To count): ReDim distances(1 To count)
    For i = 1 To count
        path.CopySegment i, edges(i): distances(i) = direction * 20#
        segmentID = path.SourceSegmentID(i)
        If segmentID = "Corner.TR" Or segmentID = "Edge.TR" Or segmentID = "Corner.TL" Then distances(i) = direction * topOffset
    Next i
    Set offset = New CGeometryVariableOffset: Set query = New CSectionGeometryQuery
    Set region = offset.Build(edges, distances, count, "CollapsedUserCorners[TopOffset]")
    region.CopyEdges edges, count
    For i = 1 To count
        If edges(i).Sweep <> 0# Then arcs = arcs + 1
        If edges(i).Sweep = 0# And Abs(edges(i).Y1 - (500# - topOffset)) < GEOM_BOUNDARY_TOLERANCE And _
                Abs(edges(i).Y2 - (500# - topOffset)) < GEOM_BOUNDARY_TOLERANCE Then
            upper = upper + 1
            TestNear "collapsed.top.intersection.left", GeomMin(edges(i).X1, edges(i).X2), 20#
            TestNear "collapsed.top.intersection.right", GeomMax(edges(i).X1, edges(i).X2), 280#
        End If
    Next i
    TestCheck "collapsed.only.bottom.arcs", arcs = 2
    TestCheck "collapsed.one.top.line", upper = 1
    TestCheck "collapsed.domain.inside", query.ContainsPoint(region, 150#, 490# - topOffset, True)
    TestCheck "collapsed.domain.outside", Not query.ContainsPoint(region, 150#, 510# - topOffset, True)
    mReport = mReport & "COLLAPSED_CORNERS: top=" & CStr(topOffset) & "; reverse=" & CStr(reverse) & "; top=(20," & CStr(500# - topOffset) & ")..(280," & CStr(500# - topOffset) & "); arcs=2" & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет контактные 2,5 мм, отключённый R2 и сдвиг SideBySide после проверки контура.
Private Sub TestContactVariants(ByVal sideBySide As Boolean, ByVal everySecond As Boolean, ByVal disabledSecond As Boolean)
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions, contact As CRebarContactBuilder
    Dim definition As CSectionDefinition, executor As CRebarRuleExecutor, layout As CRebarLayout, i As Long, count As Long
    Dim owner As Long, station As Long, expectedNormal As Double, expectedTangent As Double, binding As String
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "IndividualContact"
    Set path = facade.RectanglePath("Boundary", 0#, 0#, 1000#, 800#, "Bottom", "Right", "Top", "Left"): facade.Outer "Outer", path.ID
    facade.RebarContourBaseline "Outer", "Boundary", Array("Left", "Right"), 50#, "IndividualContact[SideLimits]"
    Set path = facade.LinePath("Top", 50#, 750#, 950#, 750#)
    Set options = New CRebarRuleOptions: options.SourceReference = "IndividualContact[Main]"
    facade.RebarByCount "Main", path.ID, 5, 10#, options
    facade.RebarContour "Main", "Outer", "Boundary", 50#, "Second", "Third"
    Set path = facade.LinePath("Left", 50#, 50#, 50#, 750#)
    facade.RebarAtStations "LeftMain", path.ID, Array(350#), 10#, options
    facade.RebarContour "LeftMain", "Outer", "Boundary", 50#, "", ""
    Set path = facade.LinePath("Right", 950#, 750#, 950#, 50#)
    facade.RebarAtStations "RightMain", path.ID, Array(350#), 10#, options
    facade.RebarContour "RightMain", "Outer", "Boundary", 50#, "", ""
    Set contact = New CRebarContactBuilder: contact.Initialize 10#, -1#, -1#, options
    binding = "EachBar": If everySecond Then binding = "EverySecondBar"
    If Not disabledSecond Then
        Set options = contact.AppendRow(10#, "Stacked", binding, "IndividualContact[Extra,2]", 40#)
        facade.RebarDependent "Second", "Main", 10#, options
    End If
    Set options = contact.AppendRow(10#, IIf(sideBySide, "SideBySide", "Stacked"), binding, "IndividualContact[Extra,3]", 40#)
    facade.RebarDependent "Third", "Main", 10#, options
    expectedNormal = 40#: expectedTangent = 0#
    If Not disabledSecond Then
        If sideBySide Then expectedTangent = -12.5 Else expectedNormal = 52.5
    End If
    Set definition = facade.Finish: Set executor = New CRebarRuleExecutor: Set layout = TestBuildLayout(definition, executor)
    For i = 1 To executor.PositionCount
        If executor.PositionRuleID(i) = "Third" Then
            count = count + 1: owner = executor.PositionPhysicalOwner(i)
            station = count - 1: If everySecond Then station = 2 * station
            TestNear "contact.normal.from.Main", layout.Y(owner), 750# - expectedNormal
            TestNear "contact.tangent.after.check", layout.X(owner), 50# + 225# * station + expectedTangent
        End If
    Next i
    TestCheck "contact.selected.count", count = IIf(everySecond, 3, 5)
    TestCheck "contact.no.false.skip", Not TestHasOffsetSkip(executor)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Объявляет оба ряда от Main с независимыми базовыми отступами и заданным контактом.
Private Sub TestAddRows(ByVal facade As CSectionDefinitionBuilder, ByVal id As String, ByVal sign As Double, _
        ByVal a2 As Double, ByVal a3 As Double, ByVal everySecond As Boolean)
    Dim contact As CRebarContactBuilder, options As CRebarRuleOptions, binding As String
    Set options = New CRebarRuleOptions: options.Coincidence = KeepPublishedOwner
    Set contact = New CRebarContactBuilder: contact.Initialize 10#, sign, 1#, options
    binding = "EachBar": If everySecond Then binding = "EverySecondBar"
    Set options = contact.AppendRow(10#, "Stacked", binding, "IndividualRectangle[Extra," & id & ",2]", a2)
    facade.RebarDependent id & ".Second", id & ".Main", 10#, options
    Set options = contact.AppendRow(10#, "Stacked", binding, "IndividualRectangle[Extra," & id & ",3]", a3)
    facade.RebarDependent id & ".Third", id & ".Main", 10#, options
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Собирает только геометрию арматуры; бетонная сетка и расчёт НДС в набор не входят.
Private Function TestBuildLayout(ByVal definition As CSectionDefinition, ByVal executor As CRebarRuleExecutor) As CRebarLayout
    Dim geometry As CGeometryRegion, region As CConcreteRegion
    Set region = definition.ConcreteDomain: Set geometry = New CGeometryRegion
    geometry.Initialize region
    Set TestBuildLayout = executor.Build(definition, geometry)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает результат проверки, сохраняя все отказы в итоговом отчёте.
Private Function TestHasOffsetSkip(ByVal executor As CRebarRuleExecutor) As Boolean
    Dim i As Long, diagnostic As CGeometryDiagnostic
    For i = 1 To executor.DiagnosticCount
        Set diagnostic = executor.DiagnosticAt(i)
        If diagnostic.Code = "OFFSET_POSITION_SKIPPED" Then
            TestHasOffsetSkip = True
            mReport = mReport & "SKIP: " & diagnostic.EntityID & "; source=" & diagnostic.RelatedEntityID & _
                "; x=" & CStr(diagnostic.Detail1) & "; y=" & CStr(diagnostic.Detail2) & "; " & diagnostic.Reason & "; " & diagnostic.SourceReference & vbCrLf
        End If
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает результат проверки, сохраняя все отказы в итоговом отчёте.
Private Sub TestCheck(ByVal name As String, ByVal success As Boolean)
    If success Then mPassed = mPassed + 1 Else mFailed = mFailed + 1
    If Not success Then mReport = mReport & "FAIL: " & name & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает координаты с допуском границы 0,000001 мм, без округления исходных значений.
Private Sub TestNear(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    TestCheck name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= GEOM_BOUNDARY_TOLERANCE
End Sub
