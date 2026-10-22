Attribute VB_Name = "modTestGeometry"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: Тесты геометрии, сетки и раскладки арматуры
' ==========================================================================
' Модуль защищает договоренности по Circle, PointedSection, HollowRectangle и Rectangles: габариты,
' дискретизацию, автоматическую арматуру и semantic-аннотации для схемы.

Private Type TTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет проверки форм, сетки, автоматической арматуры, аннотаций и
' импортированных Region fixtures. Возвращает численный протокол; наличие
' fake-импорта не объявляется проверкой настоящего сеанса AutoCAD.
Public Function RunGeometryTests() As String
    On Error GoTo Failed

    Dim stats As TTestStats

    TestRectangle stats
    TestSymmetricPointedSection stats
    TestTaperedPointedSection stats
    TestPointedSectionContourAnnotations stats
    TestPointedSectionRebarLayout stats
    TestHollowRectangleGeometry stats
    TestAudit03HollowContainedOpeningMesh stats
    TestHollowRectangleContourAnnotations stats
    TestHollowRectangleRebarLayout stats
    TestHollowRectangleSharpOpeningBProjection stats
    TestCircleGeometry stats
    TestCircleCoreDistance stats
    TestGeometryReviewCoreScale stats
    TestGeometryReviewSP35Rectangles stats
    TestAsymmetricCoreDistanceUsesOppositeFiber stats
    TestCoreDistanceWithProductInertiaMatchesLinearBoundary stats
    TestPrincipalAxisCoreDistanceIsAxisBased stats
    TestCirclePrincipalAxesStableOnCoarseMesh stats
    TestCircleInvalidData stats
    TestCircleAutoRebarLayout stats
    TestRectanglesGeometry stats
    TestRectanglesGeometryModes stats
    TestRectanglesPrincipalAxesAndCoreDistances stats
    TestConcreteCoverUsesLocalContour stats
    TestRectanglesAutoRebarLayout stats
    TestRectanglesSeparateLineOffsets stats
    TestRectanglesAdditionalRebarRows stats
    TestSectionModelFromGeneratedGeometry stats
    TestConcreteMeshCenter stats
    TestRebarAnnotationAnchors stats
    TestAutoCADImporterBuildsSectionModel stats
    TestAutoCADImporterRotatedRectangleBounds stats
    TestAutoCADImporterPrincipalInertiaBounds stats
    TestAutoCADImporterTranslatedRegionKeepsLocalGeometry stats
    TestAudit03AutoCADSmallTranslationInertia stats
    TestAutoCADImporterSquareRegionReadsEdgeRotation stats
    TestAutoCADImporterZeroEdgeAndCircleAverage stats
    TestAutoCADImporterAreaSquareFallback stats
    TestAutoCADImporterAreaSquareFallbackAverageRotation stats
    TestSectionModelEquivalentSquareEdgeRotationAverageSource stats
    TestAutoCADImporterInvalidInertiaSquareFallback stats
    TestSectionModelEquivalentRectangleFromRealInertia stats
    TestSectionModelBoundaryDoesNotReplaceRealInertia stats
    TestSectionModelFallbackInertiaSquare stats
    TestInvalidData stats
    TestBoundarySubcellMesh stats
    TestRectangularMeshSteps stats
    TestMeshConvergence stats
    TestPerformance stats
    TestAudit03GeometryLifecycle stats
    TestAudit03CircleCountGuard stats
    TestAudit03CirclePlacementAddresses stats
    TestGeometryReviewTranslationInertia stats
    TestGeometryReviewMeshFailureState stats
    TestGeometryReviewConcreteAddFailureState stats
    TestGeometryReviewLargeAngles stats
    TestGeometryReviewAreaTranslation stats
    TestGeometryReviewCornerJoints stats

    TestAppendLine stats, "TOTAL: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunGeometryTests = stats.Report
    Exit Function

Failed:
    RunGeometryTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' перенесенные геометрические fixtures без оставшихся старых builders.
Public Function RunRetiredGeometryPrimitiveTests() As String
    Dim stats As TTestStats
    On Error GoTo Failed
    TestSectionModelFromGeneratedGeometry stats
    TestRectangle stats
    TestSymmetricPointedSection stats
    TestTaperedPointedSection stats
    TestHollowRectangleGeometry stats
    TestAudit03HollowContainedOpeningMesh stats
    TestCircleGeometry stats
    TestCircleCoreDistance stats
    TestCirclePrincipalAxesStableOnCoarseMesh stats
    TestBoundarySubcellMesh stats
    TestRectangularMeshSteps stats
    TestRectanglesGeometry stats
    TestMeshConvergence stats
    TestPerformance stats
    TestAppendLine stats, "TOTAL_RETIRED_GEOMETRY_PRIMITIVES: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunRetiredGeometryPrimitiveTests = stats.Report
    Exit Function
Failed:
    RunRetiredGeometryPrimitiveTests = stats.Report & "FAIL: runtime=" & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' аналитические сопряжения и отверстия без старых shape contracts.
Public Function RunRetiredGeometryArcOpeningTests() As String
    Dim stats As TTestStats
    On Error GoTo Failed
    TestTaperedPointedSection stats
    TestHollowRectangleGeometry stats
    TestAudit03HollowContainedOpeningMesh stats
    TestAppendLine stats, "TOTAL_RETIRED_GEOMETRY_ARC_OPENING: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunRetiredGeometryArcOpeningTests = stats.Report
    Exit Function
Failed:
    RunRetiredGeometryArcOpeningTests = stats.Report & "FAIL: runtime=" & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' общий центр геометрии, численного решения и пользовательских нагрузок,
' повторная подготовка после изменений и перенос координат на 1000 м.
Public Function RunConcreteMeshCenterTests() As String
    On Error GoTo Failed
    Dim stats As TTestStats
    TestConcreteMeshCenter stats
    TestCircleCoreDistance stats
    TestAsymmetricCoreDistanceUsesOppositeFiber stats
    TestCoreDistanceWithProductInertiaMatchesLinearBoundary stats
    TestRectanglesPrincipalAxesAndCoreDistances stats
    TestAppendLine stats, "TOTAL_CONCRETE_MESH_CENTER: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunConcreteMeshCenterTests = stats.Report
    Exit Function
Failed:
    RunConcreteMeshCenterTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимый центр двух неравных площадей и третьей добавленной площади.
' Арматура не входит в веса, Clear и новая геометрия не оставляют старый кэш.
Private Sub TestConcreteMeshCenter(ByRef stats As TTestStats)
    Dim shift As Variant, section As CSectionModel, frame As CSectionNumericalFrame
    Dim props As CSectionPropertiesCalculator, centerX As Double, centerY As Double, area As Double
    Dim referenceX As Double, referenceY As Double, prefix As String, errorNumber As Long
    For Each shift In Array(0#, 1000000#, -1000000#)
        Set section = New CSectionModel
        section.AddConcreteElement CDbl(shift), -CDbl(shift), 100#, width:=10#, height:=10#
        section.AddConcreteElement CDbl(shift) + 40#, -CDbl(shift) + 80#, 300#, width:=30#, height:=10#
        prefix = "meshCenter." & CStr(shift)
        section.ConcreteMeshCentroid centerX, centerY, area
        TestAssertClose stats, prefix & ".weightedX", centerX, CDbl(shift) + 30#, 0#
        TestAssertClose stats, prefix & ".weightedY", centerY, -CDbl(shift) + 60#, 0#
        TestAssertClose stats, prefix & ".area", area, 400#, 0#
        section.ConcreteMeshCentroid centerX, centerY, area
        TestAssertClose stats, prefix & ".repeatX", centerX, CDbl(shift) + 30#, 0#
        TestAssertClose stats, prefix & ".repeatY", centerY, -CDbl(shift) + 60#, 0#
        Set frame = New CSectionNumericalFrame: frame.Initialize section
        TestAssertClose stats, prefix & ".radius", frame.LengthScale, Sqr(30# * 30# + 60# * 60#), 0#
        Set frame = New CSectionNumericalFrame: frame.Initialize section
        TestAssertClose stats, prefix & ".radiusRepeat", frame.LengthScale, Sqr(30# * 30# + 60# * 60#), 0#
        section.AddRebarElement CDbl(shift) + 500#, -CDbl(shift) - 400#, 200#, 0#, "A400"
        section.ConcreteMeshCentroid centerX, centerY, area
        TestAssertClose stats, prefix & ".rebarExcludedX", centerX, CDbl(shift) + 30#, 0#
        TestAssertClose stats, prefix & ".rebarExcludedY", centerY, -CDbl(shift) + 60#, 0#
        TestAssertClose stats, prefix & ".rebarExcludedArea", area, 400#, 0#
        Set frame = New CSectionNumericalFrame: frame.Initialize section
        Set props = New CSectionPropertiesCalculator: props.CalculateConcrete section
        CalculateConcreteSectionCentroid section, referenceX, referenceY
        TestAssertClose stats, prefix & ".frameX", frame.OriginX, centerX, 0#
        TestAssertClose stats, prefix & ".frameY", frame.OriginY, centerY, 0#
        TestAssertClose stats, prefix & ".radiusRebarInvalidation", frame.LengthScale, Sqr(470# * 470# + 460# * 460#), 0#
        TestAssertClose stats, prefix & ".propertiesX", props.CentroidX, centerX, 0#
        TestAssertClose stats, prefix & ".propertiesY", props.CentroidY, centerY, 0#
        TestAssertClose stats, prefix & ".loadReferenceX", referenceX, centerX, 0#
        TestAssertClose stats, prefix & ".loadReferenceY", referenceY, centerY, 0#
        TestAssertClose stats, prefix & ".originalX", section.ConcreteX(1), CDbl(shift), 0#
        TestAssertClose stats, prefix & ".originalY", section.ConcreteY(2), -CDbl(shift) + 80#, 0#
        section.AddConcreteElement CDbl(shift) - 20#, -CDbl(shift) - 40#, 400#
        section.ConcreteMeshCentroid centerX, centerY, area
        frame.Initialize section: props.EnsureConcrete section
        TestAssertClose stats, prefix & ".changedX", centerX, CDbl(shift) + 5#, 0#
        TestAssertClose stats, prefix & ".changedY", centerY, -CDbl(shift) + 10#, 0#
        TestAssertClose stats, prefix & ".changedArea", area, 800#, 0#
        TestAssertClose stats, prefix & ".changedFrameX", frame.OriginX, centerX, 0#
        TestAssertClose stats, prefix & ".radiusConcreteInvalidation", frame.LengthScale, Sqr(495# * 495# + 410# * 410#), 0#
        TestAssertClose stats, prefix & ".changedPropertiesY", props.CentroidY, centerY, 0#
        section.Clear
        On Error Resume Next
        section.ConcreteMeshCentroid centerX, centerY, area
        errorNumber = Err.Number: Err.Clear
        On Error GoTo 0
        TestAssertTrue stats, prefix & ".clearRejectsEmpty", errorNumber = vbObjectError + 5220
        section.AddConcreteElement CDbl(shift) + 7#, -CDbl(shift) - 9#, 25#
        section.ConcreteMeshCentroid centerX, centerY, area
        frame.Initialize section: props.EnsureConcrete section
        TestAssertClose stats, prefix & ".rebuiltX", centerX, CDbl(shift) + 7#, 0#
        TestAssertClose stats, prefix & ".rebuiltY", centerY, -CDbl(shift) - 9#, 0#
        TestAssertClose stats, prefix & ".rebuiltArea", area, 25#, 0#
        TestAssertClose stats, prefix & ".rebuiltFrameX", frame.OriginX, centerX, 0#
        TestAssertClose stats, prefix & ".rebuiltRadius", frame.LengthScale, 1#, 0#
        TestAssertClose stats, prefix & ".rebuiltPropertiesY", props.CentroidY, centerY, 0#
    Next shift
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отдельный направленный прогон локальных сумм площади: полная геометрия до WCS=1e10, сумма и знак площади также при WCS=1e12.
Public Function TestRunWorkingAreaTranslationTests() As String
    On Error GoTo Failed
    Dim stats As TTestStats
    TestGeometryReviewAreaTranslation stats
    TestAppendLine stats, "TOTAL_WORKING_AREA_TRANSLATION: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    TestRunWorkingAreaTranslationTests = stats.Report
    Exit Function
Failed:
    TestRunWorkingAreaTranslationTests = stats.Report & "FAIL: areaTranslation.runtime=" & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Полные mesh/Opening/Main/Extra пути проверяются там, где шаг Double ещё меньше геометрической пробы; экстремальные координаты отдельно проверяют площадь.
Private Sub TestGeometryReviewAreaTranslation(ByRef stats As TTestStats)
    Dim side As Variant, shift As Variant, angle As Variant, clockwise As Variant
    For Each side In Array(10#, 100#, 1000#)
        For Each shift In Array(0#, 1000000#, -1000000#, 10000000000#, -10000000000#, 1000000000000#, -1000000000000#)
            For Each clockwise In Array(False, True)
                For Each angle In Array(0#, GEOM_PI / 2#, GEOM_PI / 7#)
                    TestGeometryReviewLocalRingArea stats, CDbl(side), CDbl(shift), CDbl(angle), CBool(clockwise)
                Next angle
                TestGeometryReviewLocalArcArea stats, CDbl(side), CDbl(shift), CBool(clockwise)
            Next clockwise
        Next shift
    Next side
    For Each side In Array(10#, 100#)
        For Each shift In Array(0#, 1000000#, -1000000#, 10000000000#, -10000000000#)
            TestGeometryReviewMeshOpeningTranslation stats, CDbl(side), CDbl(shift)
            For Each clockwise In Array(False, True)
                TestGeometryReviewOffsetTranslation stats, CDbl(side), CDbl(shift), CBool(clockwise)
                TestGeometryReviewFamilyTranslation stats, CDbl(side), CDbl(shift), CBool(clockwise), False
                TestGeometryReviewFamilyTranslation stats, CDbl(side), CDbl(shift), CBool(clockwise), True
                TestGeometryReviewAnnulusTranslation stats, CDbl(side), CDbl(shift), CBool(clockwise)
            Next clockwise
        Next shift
    Next side
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит четыре точных исходных ребра квадрата, затем при необходимости обращает обход; ни один рабочий метод не получает права менять этот массив.
Private Sub TestGeometryReviewSquareEdges(ByRef edges() As TRegionEdge, ByVal side As Double, _
        ByVal shift As Double, ByVal angle As Double, ByVal clockwise As Boolean)
    Dim x(1 To 4) As Double, y(1 To 4) As Double, i As Long, j As Long, c As Double, sn As Double
    Dim reversed(1 To 4) As TRegionEdge
    x(2) = side: x(3) = side: y(3) = side: y(4) = side
    c = Cos(angle): sn = Sin(angle): ReDim edges(1 To 4)
    For i = 1 To 4
        j = i Mod 4 + 1
        edges(i).X1 = shift + x(i) * c - y(i) * sn: edges(i).Y1 = -shift + x(i) * sn + y(i) * c
        edges(i).X2 = shift + x(j) * c - y(j) * sn: edges(i).Y2 = -shift + x(j) * sn + y(j) * c
        edges(i).LoopID = 1
    Next i
    If clockwise Then
        For i = 1 To 4: reversed(i) = edges(5 - i): modGeometryPrimitives.ReverseEdge reversed(i): Next i
        For i = 1 To 4: edges(i) = reversed(i): Next i
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает локальную сумму с независимым разбиением четырёхугольника на два треугольника по фактически представленным Double-координатам.
Private Sub TestGeometryReviewLocalRingArea(ByRef stats As TTestStats, ByVal side As Double, _
        ByVal shift As Double, ByVal angle As Double, ByVal clockwise As Boolean)
    Dim edges() As TRegionEdge, variable As CGeometryVariableOffset, i As Long, area As Double, expected As Double
    Dim x0 As Double, y0 As Double, prefix As String, sourceX As Double, sourceY As Double, sign As Double
    TestGeometryReviewSquareEdges edges, side, shift, angle, clockwise
    x0 = edges(1).X1: y0 = edges(1).Y1: sourceX = edges(2).X1: sourceY = edges(2).Y1
    Set variable = New CGeometryVariableOffset
    For i = 1 To 4: area = area + variable.EdgeArea(edges(i), x0, y0): Next i
    For i = 2 To 3
        expected = expected + ((edges(i).X1 - x0) * (edges(i + 1).Y1 - y0) - (edges(i + 1).X1 - x0) * (edges(i).Y1 - y0)) / 2#
    Next i
    prefix = "review.localArea.B=" & CStr(side) & ".WCS=" & CStr(shift) & ".rot=" & CStr(angle) & ".cw=" & CStr(clockwise)
    sign = 1#: If clockwise Then sign = -1#
    TestAssertClose stats, prefix & ".triangulatedArea", area, expected, side * side * 0.000000000001
    TestAssertTrue stats, prefix & ".orientation", Sgn(area) = sign
    TestAssertTrue stats, prefix & ".sourceUnchanged", edges(1).X1 = x0 And edges(1).Y1 = y0 And edges(2).X1 = sourceX And edges(2).Y1 = sourceY
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Полная окружность из четырёх четвертей проверяет знаковую площадь хорд и дуговых сегментов, включая оба направления обхода при WCS=1e12.
Private Sub TestGeometryReviewLocalArcArea(ByRef stats As TTestStats, ByVal radius As Double, ByVal shift As Double, ByVal clockwise As Boolean)
    Dim path As CGeometryPath, variable As CGeometryVariableOffset, edge As TRegionEdge, i As Long
    Dim originX As Double, originY As Double, area As Double, sign As Double, prefix As String
    sign = 1#: If clockwise Then sign = -1#
    Set path = New CGeometryPath: path.Initialize "AreaCircle"
    For i = 0 To 3: path.AddArc "Arc" & CStr(i), shift, -shift, radius, sign * CDbl(i) * GEOM_PI / 2#, sign * GEOM_PI / 2#: Next i
    path.Seal True: path.CopySegment 1, edge: originX = edge.X1: originY = edge.Y1
    Set variable = New CGeometryVariableOffset
    For i = 1 To 4: path.CopySegment i, edge: area = area + variable.EdgeArea(edge, originX, originY): Next i
    prefix = "review.localArc.R=" & CStr(radius) & ".WCS=" & CStr(shift) & ".cw=" & CStr(clockwise)
    TestAssertClose stats, prefix & ".area", area, sign * GEOM_PI * radius * radius, radius * radius * 0.000000001
    TestAssertTrue stats, prefix & ".orientation", Sgn(area) = sign
    path.CopySegment 1, edge
    TestAssertTrue stats, prefix & ".sourceUnchanged", edge.X1 = originX And edge.Y1 = originY And edge.CenterX = shift And edge.CenterY = -shift
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Смещённый квадрат сохраняет точную площадь и обход до WCS=1e10; исходные ребра остаются в прежней системе координат.
Private Sub TestGeometryReviewOffsetTranslation(ByRef stats As TTestStats, ByVal side As Double, ByVal shift As Double, ByVal clockwise As Boolean)
    Dim edges() As TRegionEdge, distances(1 To 4) As Double, variable As CGeometryVariableOffset
    Dim region As CConcreteRegion, i As Long, sign As Double, oldX As Double, oldY As Double, prefix As String
    TestGeometryReviewSquareEdges edges, side, shift, 0#, clockwise
    oldX = edges(1).X1: oldY = edges(1).Y1: sign = 1#: If clockwise Then sign = -1#
    For i = 1 To 4: distances(i) = sign * side / 10#: Next i
    Set variable = New CGeometryVariableOffset: Set region = variable.Build(edges, distances, 4, "AreaTranslation")
    prefix = "review.offsetArea.B=" & CStr(side) & ".WCS=" & CStr(shift) & ".cw=" & CStr(clockwise)
    TestAssertClose stats, prefix & ".area", region.Area, 0.64 * side * side, side * side * 0.000000001
    TestAssertClose stats, prefix & ".oneLoop", region.LoopCount, 1#, 0#
    TestAssertTrue stats, prefix & ".sourceUnchanged", edges(1).X1 = oldX And edges(1).Y1 = oldY
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сетка с неявной наружной границей и заданным круглым отверстием проверяет локальную классификацию каждого сеточного кольца.
Private Sub TestGeometryReviewMeshOpeningTranslation(ByRef stats As TTestStats, ByVal side As Double, ByVal shift As Double)
    Dim model As CSectionModel, query As CSectionGeometryQuery, region As CConcreteRegion
    Dim i As Long, j As Long, cell As Double, prefix As String
    cell = side / 5#: Set model = New CSectionModel
    For i = 0 To 4
        For j = 0 To 4
            model.AddConcreteElement shift + (CDbl(i) + 0.5) * cell, -shift + (CDbl(j) + 0.5) * cell, cell * cell, width:=cell, height:=cell
        Next j
    Next i
    model.Contours.AddContourCircle "HOLE", shift + side / 2#, -shift + side / 2#, side / 10#, loopID:="OPENING", loopRole:="Opening"
    Set query = New CSectionGeometryQuery: query.Initialize model: Set region = query.ConcreteDomain
    prefix = "review.meshOpeningArea.B=" & CStr(side) & ".WCS=" & CStr(shift)
    TestAssertClose stats, prefix & ".area", region.Area, side * side * (1# - GEOM_PI / 100#), side * side * 0.000000001
    TestAssertClose stats, prefix & ".twoLoops", region.LoopCount, 2#, 0#
    TestAssertTrue stats, prefix & ".sourceUnchanged", model.ConcreteX(1) = shift + cell / 2# And model.ConcreteY(1) = -shift + cell / 2#
    TestAssertClose stats, prefix & ".meshCountUnchanged", model.ConcreteCount, 25#, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт Main/Extra с полной явной связью либо с Outer, либо с Opening; смена обхода меняет только знаки нормальных отступов.
Private Function TestGeometryReviewAreaFamily(ByVal side As Double, ByVal shift As Double, ByVal clockwise As Boolean, _
        ByVal opening As Boolean, ByRef executor As CRebarRuleExecutor, ByRef definition As CSectionDefinition) As CRebarLayout
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, geometry As CGeometryRegion
    Dim options As CRebarRuleOptions, edges() As TRegionEdge, i As Long, inward As Double, loopID As String
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "AreaFamily"
    TestGeometryReviewSquareEdges edges, side, shift, 0#, clockwise
    Set path = New CGeometryPath: path.Initialize "Boundary"
    For i = 1 To 4: path.AddLine "Edge" & CStr(i), edges(i).X1, edges(i).Y1, edges(i).X2, edges(i).Y2: Next i
    path.Seal True: facade.AddPath path
    inward = 1#: If clockwise Then inward = -1#
    If opening Then
        Set path = facade.RectanglePath("OuterBoundary", shift - side, -shift - side, 3# * side, 3# * side)
        facade.Outer "OUTER", path.ID
        facade.Opening "OPENING", "Boundary", "OUTER"
        loopID = "OPENING": inward = -inward
    Else
        facade.Outer "OUTER", "Boundary": loopID = "OUTER"
    End If
    Set path = facade.Offset("Axis", "Boundary", inward * side / 10#)
    Set options = New CRebarRuleOptions: options.EndMode = ClosedUniform: options.AnnotationGroup = "Family"
    facade.RebarByCount "Main", path.ID, 8, side / 20#, options
    Set options = New CRebarRuleOptions: options.NormalOffset = inward * side / 10#: options.ValidateRowOffset = True
    options.AnnotationGroup = "Family": options.PlacementCheck = CheckFullDisk
    facade.RebarDependent "Extra", "Main", side / 20#, options
    facade.RebarContour "Main", loopID, "Boundary", side / 10#, "Extra", "", normalSign:=inward
    Set definition = facade.Finish: Set geometry = New CGeometryRegion: geometry.Initialize definition.ConcreteDomain
    Set executor = New CRebarRuleExecutor: Set TestGeometryReviewAreaFamily = executor.Build(definition, geometry)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет количество, источники, владельцев и координаты фактических Main/Extra с неперенесённой моделью той же роли и обхода.
Private Sub TestGeometryReviewFamilyTranslation(ByRef stats As TTestStats, ByVal side As Double, ByVal shift As Double, _
        ByVal clockwise As Boolean, ByVal opening As Boolean)
    Dim definition As CSectionDefinition, reference As CSectionDefinition, executor As CRebarRuleExecutor, control As CRebarRuleExecutor
    Dim bars As CRebarLayout, expected As CRebarLayout, i As Long, difference As Double, sameOwners As Boolean
    Dim mainCount As Long, extraCount As Long, prefix As String, factor As Double, edge As TRegionEdge
    Set expected = TestGeometryReviewAreaFamily(side, 0#, clockwise, opening, control, reference)
    Set bars = TestGeometryReviewAreaFamily(side, shift, clockwise, opening, executor, definition)
    prefix = "review.familyArea.B=" & CStr(side) & ".WCS=" & CStr(shift) & ".cw=" & CStr(clockwise) & ".opening=" & CStr(opening)
    factor = 1#: If opening Then factor = 8#
    TestAssertClose stats, prefix & ".physicalCount", bars.Count, expected.Count, 0#
    TestAssertClose stats, prefix & ".requested", executor.RequestedPositions, 16#, 0#
    TestAssertClose stats, prefix & ".sourceArea", definition.ConcreteDomain.Area, factor * side * side, side * side * 0.000000001
    TestAssertClose stats, prefix & ".sourceLength", definition.Path("Boundary").Length, 4# * side, side * 0.000000001
    definition.Path("Boundary").CopySegment 1, edge
    TestAssertTrue stats, prefix & ".sourceUnchanged", edge.X1 = shift And edge.Y1 = -shift
    sameOwners = bars.Count = expected.Count And executor.PositionCount = control.PositionCount
    For i = 1 To GeomMin(bars.Count, expected.Count)
        difference = GeomMax(difference, Abs((bars.X(i) - shift) - expected.X(i)))
        difference = GeomMax(difference, Abs((bars.Y(i) + shift) - expected.Y(i)))
        sameOwners = sameOwners And bars.BarID(i) = expected.BarID(i) And bars.Diameter(i) = expected.Diameter(i)
    Next i
    For i = 1 To GeomMin(executor.PositionCount, control.PositionCount)
        sameOwners = sameOwners And executor.PositionRuleID(i) = control.PositionRuleID(i) And executor.PositionSourceID(i) = control.PositionSourceID(i) And executor.PositionPhysicalOwner(i) = control.PositionPhysicalOwner(i)
        If executor.PositionPhysicalOwner(i) > 0 Then
            If executor.PositionRuleID(i) = "Main" Then mainCount = mainCount + 1
            If executor.PositionRuleID(i) = "Extra" Then extraCount = extraCount + 1
        End If
    Next i
    TestAssertClose stats, prefix & ".coordinateDeviation", difference, 0#, GeomMax(0.000001, Abs(shift) * 0.0000000000000005)
    TestAssertTrue stats, prefix & ".sameOwners", sameOwners
    TestAssertClose stats, prefix & ".mainPhysical", mainCount, 8#, 0#
    If opening Then
        TestAssertClose stats, prefix & ".extraPhysical", extraCount, 8#, 0# ' У расширяющегося отверстия оба нормальных переноса угловых позиций остаются на новых прямых; у сужающегося Outer четыре угловые позиции отбрасываются.
    Else
        TestAssertClose stats, prefix & ".extraPhysical", extraCount, 4#, 0#
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Два круговых кольца с разным направлением обхода сохраняют точную площадь и одинаковую бетонную сетку при переносе до WCS=1e10.
Private Sub TestGeometryReviewAnnulusTranslation(ByRef stats As TTestStats, ByVal radius As Double, ByVal shift As Double, ByVal clockwise As Boolean)
    Dim outer As CGeometryPath, hole As CGeometryPath, query As CSectionGeometryQuery
    Dim region As CConcreteRegion, control As CConcreteRegion, geometry As CGeometryRegion, baseline As CGeometryRegion
    Dim mesh As CFiberMeshBuilder, expected As CFiberMeshBuilder, i As Long, sweep As Double, prefix As String, edge As TRegionEdge
    Set outer = New CGeometryPath: outer.Initialize "Outer"
    Set hole = New CGeometryPath: hole.Initialize "Hole"
    sweep = GEOM_PI / 2#: If clockwise Then sweep = -sweep
    For i = 0 To 3
        outer.AddArc "Outer" & CStr(i), shift, -shift, radius, CDbl(i) * sweep, sweep
        hole.AddArc "Hole" & CStr(i), shift, -shift, radius * 0.4, CDbl(i) * sweep, sweep
    Next i
    outer.Seal True: hole.Seal True: Set query = New CSectionGeometryQuery
    Set region = query.SubtractContainedRegion(query.RegionFromPath(outer), query.RegionFromPath(hole))
    Set control = query.SubtractContainedRegion(query.CircleRegion(0#, 0#, radius), query.CircleRegion(0#, 0#, radius * 0.4))
    Set geometry = New CGeometryRegion: geometry.Initialize region
    Set baseline = New CGeometryRegion: baseline.Initialize control
    Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geometry, radius / 10#, radius / 10#, 1, 2
    Set expected = New CFiberMeshBuilder: expected.BuildMesh baseline, radius / 10#, radius / 10#, 1, 2
    prefix = "review.annulusArea.R=" & CStr(radius) & ".WCS=" & CStr(shift) & ".cw=" & CStr(clockwise)
    TestAssertClose stats, prefix & ".area", region.Area, 0.84 * GEOM_PI * radius * radius, radius * radius * 0.000000001
    TestAssertClose stats, prefix & ".twoLoops", region.LoopCount, 2#, 0#
    TestAssertClose stats, prefix & ".meshCount", mesh.FiberCount, expected.FiberCount, 0#
    outer.CopySegment 1, edge
    TestAssertTrue stats, prefix & ".sourceUnchanged", edge.CenterX = shift And edge.CenterY = -shift And edge.X1 = shift + radius And edge.Y1 = -shift
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет переход generated-сетки и стержней в общую CSectionModel:
' площади/число элементов сохраняются, расчетные ID создаются моделью,
' исходное имя стержня остается отдельным трассировочным полем.
Private Sub TestSectionModelFromGeneratedGeometry(ByRef stats As TTestStats)
    Dim geom As CGeometryRegion
    Set geom = TestFixtureDisk(300#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 50#, 50#, 1, 2

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "CircleSourceName-1", 100#, 0#, 20#, 0#, "A400", "source name", geom
    rebars.AddBar "CircleSourceName-2", -100#, 0#, 20#, 0#, "A400", "source name", geom

    Dim model As CSectionModel
    Set model = TestBuildGeneratedSectionModel(mesh, rebars, "Circle")

    TestAssertTrue stats, "model.concrete.count", model.ConcreteCount = mesh.FiberCount
    TestAssertTrue stats, "model.rebar.count", model.RebarCount = rebars.Count
    TestAssertTrue stats, "model.concrete.id", model.ConcreteID(1) = "C1"
    TestAssertTrue stats, "model.rebar.id", model.RebarID(1) = "R1"
    TestAssertTrue stats, "model.rebar.sourceName", model.RebarSourceName(1) = "CircleSourceName-1"
    TestAssertClose stats, "model.rebar.area", model.RebarArea(1), GEOM_PI * 20# * 20# / 4#, 0.000000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет точки/нормали аннотаций круговых и гранных рядов арматуры.
' Нулевая активная линия не должна давать подпись отсутствующих стержней;
' перенос аннотаций в модель не меняет расположение самих стержней.
Private Sub TestRebarAnnotationAnchors(ByRef stats As TTestStats)
    Dim circleBuilder As CRebarRuleExecutor
    Set circleBuilder = New CRebarRuleExecutor

    Dim circleLabels As CSectionAnnotations, rectanglesLabels As CSectionAnnotations
    Dim zeroFaceLabels As CSectionAnnotations, zeroCircleLabels As CSectionAnnotations
    Dim circleBars As CRebarLayout
    Set circleBars = TestFixtureAnnotatedLayout(TestFixtureCircleDefinition(Array(300#, 40#, 8, 20#)), circleBuilder, circleLabels)
    Set circleBars = TestFixtureMoveLayout(circleBars, 10#, -20#)
    TestFixtureMoveAnnotations circleLabels, 10#, -20#
    TestAssertTrue stats, "annotation.circle.count", circleLabels.Count = 1
    TestAssertTrue stats, "annotation.circle.group", circleLabels.AnnotationID(1) = "REBAR_Circle"
    TestAssertClose stats, "annotation.circle.startX", circleLabels.StartX(1), -100#, 0.000001
    TestAssertClose stats, "annotation.circle.endX", circleLabels.EndX(1), 120#, 0.000001
    TestAssertClose stats, "annotation.circle.normalY", circleLabels.NormalY(1), 1#, 0.000001
    TestAssertClose stats, "annotation.circle.axisDistance", 150# - Sqr((circleBars.X(1) - 10#) ^ 2 + (circleBars.Y(1) + 20#) ^ 2), 40#, 0.000001

    Dim rectanglesBuilder As CRebarRuleExecutor
    Set rectanglesBuilder = New CRebarRuleExecutor
    Dim rectanglesBars As CRebarLayout
    Set rectanglesBars = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(Array(40#, 40#, 20#, 20#, 3, 0, 50#, 70#, 0#, 0#, 0#, 0#, 0#, 0#, "Stacked", "Stacked"), TestEmptyFaceSettings(), TestEmptyFaceSettings(), TestEmptyFaceSettings())), rectanglesBuilder, rectanglesLabels)

    TestAssertTrue stats, "annotation.rectangles.count", rectanglesLabels.Count = 1
    TestAssertTrue stats, "annotation.rectangles.group", rectanglesLabels.AnnotationID(1) = "REBAR_H1.as_1"
    TestAssertClose stats, "annotation.rectangles.startY", rectanglesLabels.StartY(1), 300#, 0.000001
    TestAssertClose stats, "annotation.rectangles.endY", rectanglesLabels.EndY(1), 730#, 0.000001
    TestAssertClose stats, "annotation.rectangles.normalX", rectanglesLabels.NormalX(1), -1#, 0.000001
    TestAssertClose stats, "annotation.rectangles.axisDistance", rectanglesBars.X(1), 40#, 0.000001

    Dim noBars As CRebarLayout
    Set noBars = New CRebarLayout
    Dim synthetic(1 To 1) As TSectionAnnotation, noBarLabels As CSectionAnnotations, formatter As CRebarGroupAnnotationBuilder
    synthetic(1).Name = "SyntheticOnly": synthetic(1).EndX = 1#: synthetic(1).NormalY = 1#
    Set noBarLabels = New CSectionAnnotations: Set formatter = New CRebarGroupAnnotationBuilder
    formatter.AddRebarLabels noBarLabels, noBars, synthetic, 1
    TestAssertTrue stats, "annotation.manual.noPhysicalLabel", noBarLabels.Count = 0

    Dim zeroFaceBuilder As CRebarRuleExecutor
    Set zeroFaceBuilder = New CRebarRuleExecutor
    Dim zeroFaceBars As CRebarLayout
    Set zeroFaceBars = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(TestEmptyFaceSettings(), TestEmptyFaceSettings(), Array(40#, 40#, 20#, 20#, 0, 0, 50#, 70#, 0#, 0#, 0#, 0#, 0#, 0#, "Stacked", "Stacked"), Array(40#, 40#, 20#, 20#, 2, 0, 50#, 70#, 0#, 0#, 0#, 0#, 0#, 0#, "Stacked", "Stacked"))), zeroFaceBuilder, zeroFaceLabels)
    TestAssertTrue stats, "annotation.zeroFace.hasOtherBars", zeroFaceBars.Count = 2
    TestAssertTrue stats, "annotation.zeroFace.noAnchor", Not TestHasAnnotationGroup(zeroFaceLabels, "B1.as_1")
    TestAssertTrue stats, "annotation.zeroFace.otherAnchor", TestHasAnnotationGroup(zeroFaceLabels, "B2.as_1")

    Dim zeroCircleBars As CRebarLayout
    Set zeroCircleBars = TestFixtureMoveLayout(TestFixtureAnnotatedLayout(TestFixtureCircleDefinition(Array(300#, 40#, 0, 20#, 20#, 30#, "Stacked", "Stacked")), circleBuilder, zeroCircleLabels), 10#, -20#)
    TestAssertTrue stats, "annotation.circle.zeroCount.noBars", zeroCircleBars.Count = 0
    TestAssertTrue stats, "annotation.circle.zeroCount.noAnchor", zeroCircleLabels.Count = 0

    Dim geom As CGeometryRegion
    Set geom = TestFixtureDisk(300#, 10#, -20#)
    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 50#, 50#, 1, 2

    Dim model As CSectionModel
    Set model = TestBuildGeneratedSectionModel(mesh, circleBars, "Circle")
    Dim annotationBuilder As CSectionAnnotationBuilder
    Set annotationBuilder = New CSectionAnnotationBuilder
    model.Contours.AddContourCircle "CONTOUR_CIRCLE", 10#, -20#, 150#, loopID:="OUTER", loopRole:="Outer"
    annotationBuilder.Build TestFixtureCircleDefinition(Array(300#, 40#, 8, 20#, 20#, 30#, "Stacked", "SideBySide")), _
        circleBars, model.Annotations, circleBuilder
    TestAssertTrue stats, "annotation.model.count", model.AnnotationCount = 2 And model.Contours.Count = 1
    TestAssertTrue stats, "annotation.model.rebarLabel", TestHasSectionAnnotation(model, "REBAR_ANNOTATION", "REBAR_Circle")
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Импортирует заданные массивы Region без AutoCAD COM: отбрасывает малую
' область, сохраняет реальные размеры/площадь и создает расчетные ID.
' Диаметр арматуры по площади и габарит бетона проверяются независимо.
Private Sub TestAutoCADImporterBuildsSectionModel(ByRef stats As TTestStats)
    Dim concreteRegions(1 To 2, 1 To 10) As Variant
    concreteRegions(1, 1) = 0.0000000001
    concreteRegions(1, 2) = 0#
    concreteRegions(1, 3) = 0#
    concreteRegions(2, 1) = 18#
    concreteRegions(2, 2) = 10#
    concreteRegions(2, 3) = 20#
    concreteRegions(2, 4) = 0#
    concreteRegions(2, 5) = 0#
    concreteRegions(2, 6) = 0#
    concreteRegions(2, 7) = "ABC"
    concreteRegions(2, 8) = 6#
    concreteRegions(2, 9) = 3#
    concreteRegions(2, 10) = 0#

    Dim rebarRegions(1 To 1, 1 To 4) As Variant
    rebarRegions(1, 1) = GEOM_PI * 20# * 20# / 4#
    rebarRegions(1, 2) = -30#
    rebarRegions(1, 3) = 40#
    rebarRegions(1, 4) = "DEF"

    Dim importer As CAutoCADSectionModelImporter
    Set importer = New CAutoCADSectionModelImporter

    Dim model As CSectionModel
    Set model = importer.TestBuildFromRegionArrays(concreteRegions, rebarRegions, "Rebar", 0.000001)

    TestAssertTrue stats, "autocad.import.source", model.SourceType = "AutoCADImport"
    TestAssertTrue stats, "autocad.import.concrete.count", model.ConcreteCount = 1
    TestAssertTrue stats, "autocad.import.rebar.count", model.RebarCount = 1
    TestAssertTrue stats, "autocad.import.concrete.id", model.ConcreteID(1) = "C1"
    TestAssertClose stats, "autocad.import.concrete.x", model.ConcreteX(1), 10#, 0.000001
    TestAssertClose stats, "autocad.import.concrete.y", model.ConcreteY(1), 20#, 0.000001
    TestAssertClose stats, "autocad.import.concrete.width", model.ConcreteWidth(1), 6#, 0.000001
    TestAssertClose stats, "autocad.import.concrete.height", model.ConcreteHeight(1), 3#, 0.000001
    TestAssertClose stats, "autocad.import.concrete.rotation", model.ConcreteRotation(1), 0#, 0.000001
    TestAssertClose stats, "autocad.import.concrete.localIx", model.ConcreteLocalIx(1), 18# * 3# * 3# / 12#, 0.000001
    TestAssertClose stats, "autocad.import.concrete.localIy", model.ConcreteLocalIy(1), 18# * 6# * 6# / 12#, 0.000001
    TestAssertClose stats, "autocad.import.rebar.diameter", model.RebarDiameter(1), 20#, 0.000001
    TestAssertTrue stats, "autocad.import.rebar.marker", model.RebarSteelClass(1) = "Rebar"
    TestAssertTrue stats, "autocad.import.rebar.comment", _
        InStr(1, model.RebarComment(1), "по площади", vbTextCompare) > 0

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    Dim minProjection As Double
    Dim maxProjection As Double
    props.CalculateProjection model, 1#, 0#, False, minProjection, maxProjection
    TestAssertClose stats, "autocad.import.concrete.xDepth", maxProjection - minProjection, 6#, 0.000001
    props.CalculateProjection model, 0#, 1#, False, minProjection, maxProjection
    TestAssertClose stats, "autocad.import.concrete.yDepth", maxProjection - minProjection, 3#, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет сохранение настоящих ширины, высоты и угла повёрнутого
' прямоугольного элемента. Габариты должны определяться по его форме,
' а не по заменяющему квадрату той же площади.
Private Sub TestAutoCADImporterRotatedRectangleBounds(ByRef stats As TTestStats)
    Dim widthValue As Double
    Dim heightValue As Double
    Dim angle As Double
    widthValue = 80#
    heightValue = 30#
    angle = GEOM_PI / 4#

    Dim concreteRegions(1 To 1, 1 To 10) As Variant
    concreteRegions(1, 1) = widthValue * heightValue
    concreteRegions(1, 2) = 0#
    concreteRegions(1, 3) = 0#
    concreteRegions(1, 8) = widthValue
    concreteRegions(1, 9) = heightValue
    concreteRegions(1, 10) = angle

    Dim rebarRegions(1 To 1, 1 To 4) As Variant
    rebarRegions(1, 1) = GEOM_PI * 12# * 12# / 4#
    rebarRegions(1, 2) = 0#
    rebarRegions(1, 3) = 0#
    rebarRegions(1, 4) = "RB"

    Dim importer As CAutoCADSectionModelImporter
    Set importer = New CAutoCADSectionModelImporter
    Dim model As CSectionModel
    Set model = importer.TestBuildFromRegionArrays(concreteRegions, rebarRegions, "Rebar", 0.000001)

    TestAssertTrue stats, "autocad.import.rotated.shape", model.ConcreteShapeType(1) = "Rectangle"
    TestAssertClose stats, "autocad.import.rotated.width", model.ConcreteWidth(1), widthValue, 0.000001
    TestAssertClose stats, "autocad.import.rotated.height", model.ConcreteHeight(1), heightValue, 0.000001
    TestAssertClose stats, "autocad.import.rotated.rotation", model.ConcreteRotation(1), angle, 0.000001
    Dim ixLocal As Double
    Dim iyLocal As Double
    Dim c As Double
    Dim s As Double
    ixLocal = concreteRegions(1, 1) * heightValue * heightValue / 12#
    iyLocal = concreteRegions(1, 1) * widthValue * widthValue / 12#
    c = Cos(angle)
    s = Sin(angle)
    TestAssertClose stats, "autocad.import.rotated.localIx", model.ConcreteLocalIx(1), _
        ixLocal * c * c + iyLocal * s * s, 0.000001
    TestAssertClose stats, "autocad.import.rotated.localIy", model.ConcreteLocalIy(1), _
        ixLocal * s * s + iyLocal * c * c, 0.000001
    TestAssertClose stats, "autocad.import.rotated.localIxy", model.ConcreteLocalIxy(1), _
        (iyLocal - ixLocal) * s * c, 0.000001

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    Dim minProjection As Double
    Dim maxProjection As Double
    props.CalculateProjection model, Cos(angle), Sin(angle), False, minProjection, maxProjection
    TestAssertClose stats, "autocad.import.rotated.uDepth", maxProjection - minProjection, widthValue, 0.000001
    props.CalculateProjection model, -Sin(angle), Cos(angle), False, minProjection, maxProjection
    TestAssertClose stats, "autocad.import.rotated.vDepth", maxProjection - minProjection, heightValue, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что импортированный Region с реальными A/I настоящего
' прямоугольника получает оболочку Rectangle из CSectionModel. Importer при
' этом не передает Width/Height и не обращается к Explode.
Private Sub TestAutoCADImporterPrincipalInertiaBounds(ByRef stats As TTestStats)
    Dim area As Double
    Dim expectedWidth As Double
    Dim expectedHeight As Double
    Dim angle As Double
    area = 2000#
    expectedWidth = 100#
    expectedHeight = 20#
    angle = GEOM_PI / 6#

    Dim inertiaU As Double
    Dim inertiaV As Double
    Dim localIx As Double
    Dim localIy As Double
    Dim localIxy As Double
    inertiaU = area * expectedHeight * expectedHeight / 12#
    inertiaV = area * expectedWidth * expectedWidth / 12#
    TestRotatedLocalInertia inertiaU, inertiaV, angle, localIx, localIy, localIxy

    Dim concreteRegions(1 To 1, 1 To 7) As Variant
    concreteRegions(1, 1) = area
    concreteRegions(1, 2) = 0#
    concreteRegions(1, 3) = 0#
    concreteRegions(1, 4) = localIx
    concreteRegions(1, 5) = localIy
    concreteRegions(1, 6) = localIxy
    concreteRegions(1, 7) = "ROT"

    Dim rebarRegions(1 To 1, 1 To 4) As Variant
    rebarRegions(1, 1) = GEOM_PI * 12# * 12# / 4#
    rebarRegions(1, 2) = 0#
    rebarRegions(1, 3) = 0#
    rebarRegions(1, 4) = "RB"

    Dim importer As CAutoCADSectionModelImporter
    Set importer = New CAutoCADSectionModelImporter
    Dim model As CSectionModel
    Set model = importer.TestBuildFromRegionArrays(concreteRegions, rebarRegions, "Rebar", 0.000001)

    TestAssertTrue stats, "autocad.import.principal.shape", model.ConcreteShapeType(1) = "Rectangle"
    TestAssertClose stats, "autocad.import.principal.localIx", model.ConcreteLocalIx(1), localIx, 0.000001
    TestAssertClose stats, "autocad.import.principal.localIy", model.ConcreteLocalIy(1), localIy, 0.000001
    TestAssertClose stats, "autocad.import.principal.localIxy", model.ConcreteLocalIxy(1), localIxy, 0.000001
    TestAssertClose stats, "autocad.import.principal.areaShell", _
        model.ConcreteWidth(1) * model.ConcreteHeight(1), area, 0.000001
    TestAssertClose stats, "autocad.import.principal.noExplode", importer.TestDebugEdgeProbeCount, 0#, 0.000001

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    Dim minProjection As Double
    Dim maxProjection As Double
    props.CalculateProjection model, Cos(angle), Sin(angle), False, minProjection, maxProjection
    TestAssertClose stats, "autocad.import.principal.uDepth", maxProjection - minProjection, expectedWidth, 0.000001
    props.CalculateProjection model, -Sin(angle), Cos(angle), False, minProjection, maxProjection
    TestAssertClose stats, "autocad.import.principal.vDepth", maxProjection - minProjection, expectedHeight, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что перенос одного и того же повернутого Region по чертежу
' не меняет его центральные A/I и геометрическую оболочку.
Private Sub TestAutoCADImporterTranslatedRegionKeepsLocalGeometry(ByRef stats As TTestStats)
    Dim area As Double
    Dim expectedWidth As Double
    Dim expectedHeight As Double
    Dim angle As Double
    area = 2250#
    expectedWidth = 90#
    expectedHeight = 25#
    angle = -GEOM_PI / 9#

    Dim inertiaU As Double
    Dim inertiaV As Double
    Dim localIx As Double
    Dim localIy As Double
    Dim localIxy As Double
    inertiaU = area * expectedHeight * expectedHeight / 12#
    inertiaV = area * expectedWidth * expectedWidth / 12#
    TestRotatedLocalInertia inertiaU, inertiaV, angle, localIx, localIy, localIxy

    Dim modelSpace As Collection
    Set modelSpace = New Collection
    modelSpace.Add TestFakeRegion(area, 0#, 0#, localIx, localIy, localIxy, _
        "Concrete", "C_NEAR", angle, False)
    modelSpace.Add TestFakeRegion(area, 12000#, -8000#, localIx, localIy, localIxy, _
        "Concrete", "C_FAR", angle, False)
    modelSpace.Add TestFakeRegion(GEOM_PI * 12# * 12# / 4#, 0#, -60#, 1#, 1#, 0#, _
        "Reinf", "R1", 0#, False)

    Dim importer As CAutoCADSectionModelImporter
    Set importer = New CAutoCADSectionModelImporter
    Dim model As CSectionModel
    Set model = importer.ImportFromModelSpace(modelSpace, "Concrete", "Reinf", "Rebar", 0.000001)

    TestAssertClose stats, "autocad.import.translated.noExplode", importer.TestDebugEdgeProbeCount, 0#, 0.000001
    TestAssertTrue stats, "autocad.import.translated.nearShape", model.ConcreteShapeType(1) = "Rectangle"
    TestAssertTrue stats, "autocad.import.translated.farShape", model.ConcreteShapeType(2) = model.ConcreteShapeType(1)
    TestAssertClose stats, "autocad.import.translated.localIx", model.ConcreteLocalIx(2), model.ConcreteLocalIx(1), 0.001
    TestAssertClose stats, "autocad.import.translated.localIy", model.ConcreteLocalIy(2), model.ConcreteLocalIy(1), 0.001
    TestAssertClose stats, "autocad.import.translated.localIxy", model.ConcreteLocalIxy(2), model.ConcreteLocalIxy(1), 0.001
    TestAssertClose stats, "autocad.import.translated.width", model.ConcreteWidth(2), model.ConcreteWidth(1), 0.000001
    TestAssertClose stats, "autocad.import.translated.height", model.ConcreteHeight(2), model.ConcreteHeight(1), 0.000001
    TestAssertClose stats, "autocad.import.translated.rotation", model.ConcreteRotation(2), model.ConcreteRotation(1), 0.000001

    Dim nearI1 As Double
    Dim nearI2 As Double
    Dim nearAngle As Double
    Dim farI1 As Double
    Dim farI2 As Double
    Dim farAngle As Double
    model.PrincipalLocalInertia model.ConcreteLocalIx(1), model.ConcreteLocalIy(1), _
        model.ConcreteLocalIxy(1), nearI1, nearI2, nearAngle
    model.PrincipalLocalInertia model.ConcreteLocalIx(2), model.ConcreteLocalIy(2), _
        model.ConcreteLocalIxy(2), farI1, farI2, farAngle
    TestAssertClose stats, "autocad.import.translated.i1", farI1, nearI1, 0.001
    TestAssertClose stats, "autocad.import.translated.i2", farI2, nearI2, 0.001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет центральный тензор при малом переносе повернутого
' прямоугольника. При таком переносе несколько ошибочных Ixy также дают
' положительную матрицу, поэтому проверка только ее определенности недостаточна.
Private Sub TestAudit03AutoCADSmallTranslationInertia(ByRef stats As TTestStats)
    Dim signValue As Long, position As Long, angle As Double, cx As Double, cy As Double
    Dim ix As Double, iy As Double, ixy As Double, space As Collection
    Dim importer As CAutoCADSectionModelImporter, model As CSectionModel, prefix As String
    Set importer = New CAutoCADSectionModelImporter
    For signValue = -1 To 1 Step 2
        angle = CDbl(signValue) * GEOM_PI / 6#
        TestRotatedLocalInertia 2160000#, 8640000#, angle, ix, iy, ixy
        For position = 0 To 2
            cx = 0#: cy = 0#
            If position > 0 Then cx = 20#: cy = -10#
            If position = 2 Then cy = 10#
            Set space = New Collection
            space.Add TestFakeRegion(7200#, cx, cy, ix, iy, ixy, "Concrete", "C1")
            space.Add TestFakeRegion(100#, 0#, 0#, 1#, 1#, 0#, "Reinf", "R1")
            Set model = importer.ImportFromModelSpace(space, "Concrete", "Reinf", "A400", 0#)
            prefix = "audit03.autocad.smallTranslation." & CStr(signValue) & "." & CStr(position)
            TestAssertClose stats, prefix & ".area", model.ConcreteArea(1), 7200#, 0.000001
            TestAssertClose stats, prefix & ".x", model.ConcreteX(1), cx, 0.000001
            TestAssertClose stats, prefix & ".y", model.ConcreteY(1), cy, 0.000001
            TestAssertClose stats, prefix & ".Ix", model.ConcreteLocalIx(1), ix, 0.000001
            TestAssertClose stats, prefix & ".Iy", model.ConcreteLocalIy(1), iy, 0.000001
            TestAssertClose stats, prefix & ".Ixy", model.ConcreteLocalIxy(1), ixy, 0.000001
        Next position
    Next signValue
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет живой путь AutoCAD-import: почти изотропный Region получает угол
' простой грани через дешевый Explode, а неизотропный Region обходится без
' Explode и интерпретируется по своим A/I уже в CSectionModel.
Private Sub TestAutoCADImporterSquareRegionReadsEdgeRotation(ByRef stats As TTestStats)
    Dim modelSpace As Collection
    Set modelSpace = New Collection

    Dim rectArea As Double
    Dim rectWidth As Double
    Dim rectHeight As Double
    Dim rectAngle As Double
    Dim rectIx As Double
    Dim rectIy As Double
    Dim rectIxy As Double
    rectWidth = 80#
    rectHeight = 30#
    rectArea = rectWidth * rectHeight
    rectAngle = GEOM_PI / 7#
    TestRotatedLocalInertia rectArea * rectHeight * rectHeight / 12#, _
        rectArea * rectWidth * rectWidth / 12#, rectAngle, rectIx, rectIy, rectIxy
    modelSpace.Add TestFakeRegion(rectArea, 0#, 0#, rectIx, rectIy, rectIxy, _
        "Concrete", "C_RECT", rectAngle, True)

    Dim squareArea As Double
    Dim squareAngle As Double
    squareArea = 2500#
    squareAngle = GEOM_PI / 5#
    modelSpace.Add TestFakeRegion(squareArea, 120#, 0#, squareArea * squareArea / 12#, _
        squareArea * squareArea / 12#, 0#, "Concrete", "C_SQ", squareAngle, True)

    modelSpace.Add TestFakeRegion(GEOM_PI * 12# * 12# / 4#, 0#, -60#, 1#, 1#, 0#, _
        "Reinf", "R1", 0#, False)

    Dim importer As CAutoCADSectionModelImporter
    Set importer = New CAutoCADSectionModelImporter
    Dim model As CSectionModel
    Set model = importer.ImportFromModelSpace(modelSpace, "Concrete", "Reinf", "Rebar", 0.000001)

    TestAssertTrue stats, "autocad.import.edgeProbe.rect.shape", model.ConcreteShapeType(1) = "Rectangle"
    TestAssertClose stats, "autocad.import.edgeProbe.rect.noExplode", importer.TestDebugEdgeProbeCount, 1#, 0.000001
    TestAssertTrue stats, "autocad.import.edgeProbe.square.shape", model.ConcreteShapeType(2) = "Equivalent square"
    TestAssertClose stats, "autocad.import.edgeProbe.square.width", model.ConcreteWidth(2), Sqr(squareArea), 0.000001
    TestAssertClose stats, "autocad.import.edgeProbe.square.height", model.ConcreteHeight(2), Sqr(squareArea), 0.000001
    TestAssertClose stats, "autocad.import.edgeProbe.square.rotation", model.ConcreteRotation(2), squareAngle, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет известный нулевой угол прямой грани отдельно от неизвестного
' угла круга. Среднее берется по двум исходным ориентированным элементам;
' повторное назначение не должно включать круг как новый источник ориентации.
Private Sub TestAutoCADImporterZeroEdgeAndCircleAverage(ByRef stats As TTestStats)
    Dim modelSpace As Collection, importer As CAutoCADSectionModelImporter, model As CSectionModel
    Dim ix As Double, iy As Double, ixy As Double, area As Double
    Set modelSpace = New Collection
    area = 2500#
    modelSpace.Add TestFakeRegion(area, 0#, 0#, area * area / 12#, area * area / 12#, 0#, _
        "Concrete", "SQUARE_0", 0#, True)
    area = 80# * 160#
    TestRotatedLocalInertia area * 160# ^ 2 / 12#, area * 80# ^ 2 / 12#, GEOM_PI / 6#, ix, iy, ixy
    modelSpace.Add TestFakeRegion(area, 200#, 0#, ix, iy, ixy, "Concrete", "RECT_30", 0#, False)
    area = GEOM_PI * 20# ^ 2
    modelSpace.Add TestFakeRegion(area, 400#, 0#, GEOM_PI * 20# ^ 4 / 4#, GEOM_PI * 20# ^ 4 / 4#, 0#, _
        "Concrete", "CIRCLE", 0#, False)
    modelSpace.Add TestFakeRegion(100#, 0#, 0#, 1#, 1#, 0#, "Reinf", "R1", 0#, False)
    Set importer = New CAutoCADSectionModelImporter
    Set model = importer.ImportFromModelSpace(modelSpace, "Concrete", "Reinf", "A400", 0#)
    TestAssertClose stats, "autocad.import.zeroEdge.square", model.ConcreteRotation(1), 0#, 0.000001
    TestAssertClose stats, "autocad.import.zeroEdge.mean", model.AverageKnownConcreteElementRotation(), GEOM_PI / 12#, 0.000001
    TestAssertClose stats, "autocad.import.zeroEdge.circle", model.ConcreteRotation(3), GEOM_PI / 12#, 0.000001
    model.ApplyAverageRotationToEquivalentAreaFallbacks
    model.ApplyAverageRotationToEquivalentAreaFallbacks
    TestAssertClose stats, "autocad.import.zeroEdge.repeated.square", model.ConcreteRotation(1), 0#, 0.000001
    TestAssertClose stats, "autocad.import.zeroEdge.repeated.circle", model.ConcreteRotation(3), GEOM_PI / 12#, 0.000001
    TestAssertClose stats, "autocad.import.zeroEdge.repeated.mean", model.AverageKnownConcreteElementRotation(), GEOM_PI / 12#, 0.000001
    model.AddConcreteElement 600#, 0#, 80# * 160#, 1, "new rectangle", "RECT_MINUS_30", _
        "Rectangle", 80#, 160#, -GEOM_PI / 6#
    model.ApplyAverageRotationToEquivalentAreaFallbacks
    TestAssertClose stats, "autocad.import.zeroEdge.changed.square", model.ConcreteRotation(1), 0#, 0.000001
    TestAssertClose stats, "autocad.import.zeroEdge.changed.circle", model.ConcreteRotation(3), 0#, 0.000001
    TestAssertClose stats, "autocad.import.zeroEdge.changed.mean", model.AverageKnownConcreteElementRotation(), 0#, 0.000001
    TestAssertClose stats, "autocad.import.zeroEdge.changed.circleIx", model.ConcreteLocalIx(3), GEOM_PI * 20# ^ 4 / 4#, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Если у Region нет местных моментов инерции, его габариты задаёт квадрат
' той же площади. Круг радиусом по площади здесь не используется.
Private Sub TestAutoCADImporterAreaSquareFallback(ByRef stats As TTestStats)
    Dim concreteRegions(1 To 1, 1 To 7) As Variant
    concreteRegions(1, 1) = GEOM_PI * 25# * 25#
    concreteRegions(1, 2) = 0#
    concreteRegions(1, 3) = 0#
    concreteRegions(1, 7) = "AREA_ONLY"

    Dim rebarRegions(1 To 1, 1 To 4) As Variant
    rebarRegions(1, 1) = GEOM_PI * 12# * 12# / 4#
    rebarRegions(1, 2) = 0#
    rebarRegions(1, 3) = 0#
    rebarRegions(1, 4) = "RB"

    Dim importer As CAutoCADSectionModelImporter
    Set importer = New CAutoCADSectionModelImporter
    Dim model As CSectionModel
    Set model = importer.TestBuildFromRegionArrays(concreteRegions, rebarRegions, "Rebar", 0.000001)

    TestAssertTrue stats, "autocad.import.squareFallback.shape", model.ConcreteShapeType(1) = "Equivalent square"
    TestAssertClose stats, "autocad.import.squareFallback.width", model.ConcreteWidth(1), Sqr(model.ConcreteArea(1)), 0.000001
    TestAssertClose stats, "autocad.import.squareFallback.height", model.ConcreteHeight(1), Sqr(model.ConcreteArea(1)), 0.000001
    TestAssertClose stats, "autocad.import.squareFallback.localIx", model.ConcreteLocalIx(1), _
        model.ConcreteArea(1) * model.ConcreteArea(1) / 12#, 0.000001

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    Dim minProjection As Double
    Dim maxProjection As Double
    props.CalculateProjection model, 1#, 0#, False, minProjection, maxProjection
    TestAssertClose stats, "autocad.import.squareFallback.xDepth", maxProjection - minProjection, Sqr(model.ConcreteArea(1)), 0.000001
    props.CalculateProjection model, 1#, 1#, False, minProjection, maxProjection
    TestAssertClose stats, "autocad.import.squareFallback.diagonalDepth", maxProjection - minProjection, _
        Sqr(model.ConcreteArea(1)) * Sqr(2#), 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Квадрат той же площади получает средний угол распознанной сетки. Проверяет
' согласованность расчётных проекций, схемы Excel и экспорта AutoCAD.
Private Sub TestAutoCADImporterAreaSquareFallbackAverageRotation(ByRef stats As TTestStats)
    Dim widthValue As Double
    Dim heightValue As Double
    Dim angle As Double
    widthValue = 80#
    heightValue = 40#
    angle = GEOM_PI / 6#

    Dim concreteRegions(1 To 2, 1 To 10) As Variant
    concreteRegions(1, 1) = widthValue * heightValue
    concreteRegions(1, 2) = 0#
    concreteRegions(1, 3) = 0#
    concreteRegions(1, 8) = widthValue
    concreteRegions(1, 9) = heightValue
    concreteRegions(1, 10) = angle
    concreteRegions(2, 1) = 2500#
    concreteRegions(2, 2) = 200#
    concreteRegions(2, 3) = 0#
    concreteRegions(2, 7) = "AREA_ONLY"

    Dim rebarRegions(1 To 1, 1 To 4) As Variant
    rebarRegions(1, 1) = GEOM_PI * 12# * 12# / 4#
    rebarRegions(1, 2) = 0#
    rebarRegions(1, 3) = 0#
    rebarRegions(1, 4) = "RB"

    Dim importer As CAutoCADSectionModelImporter
    Set importer = New CAutoCADSectionModelImporter
    Dim model As CSectionModel
    Set model = importer.TestBuildFromRegionArrays(concreteRegions, rebarRegions, "Rebar", 0.000001)

    TestAssertTrue stats, "autocad.import.squareFallbackRotation.shape", model.ConcreteShapeType(2) = "Equivalent square"
    TestAssertClose stats, "autocad.import.squareFallbackRotation.angle", model.ConcreteRotation(2), angle, 0.000001

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    Dim minProjection As Double
    Dim maxProjection As Double
    props.CalculateProjection model, Cos(angle), Sin(angle), False, minProjection, maxProjection
    Dim expectedProjectionDepth As Double
    expectedProjectionDepth = 200# * Cos(angle) + Sqr(model.ConcreteArea(2)) / 2# + widthValue / 2#
    TestAssertClose stats, "autocad.import.squareFallbackRotation.projectedSide", _
        maxProjection - minProjection, expectedProjectionDepth, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что почти изотропный Region с уже найденным углом простой грани
' участвует в средней ориентации сетки. Такой элемент сам остается
' Equivalent square, но его ненулевой Rotation является полезной подсказкой
' для соседних элементов, у которых угол не удалось определить.
Private Sub TestSectionModelEquivalentSquareEdgeRotationAverageSource(ByRef stats As TTestStats)
    Dim areaKnown As Double
    Dim areaUnknown As Double
    Dim edgeAngle As Double
    areaKnown = 2500#
    areaUnknown = 1600#
    edgeAngle = GEOM_PI / 7#

    Dim model As CSectionModel
    Set model = New CSectionModel
    model.AddConcreteElement 0#, 0#, areaKnown, 1, "square", "S1", "Region", _
        0#, 0#, edgeAngle, vbNullString, areaKnown * areaKnown / 12#, _
        areaKnown * areaKnown / 12#, 0#
    model.AddConcreteElement 100#, 0#, areaUnknown, 1, "unknown", "S2", "Region"

    model.ApplyAverageRotationToEquivalentAreaFallbacks

    TestAssertTrue stats, "model.squareEdgeAverage.source.shape", _
        model.ConcreteShapeType(1) = "Equivalent square"
    TestAssertTrue stats, "model.squareEdgeAverage.target.shape", _
        model.ConcreteShapeType(2) = "Equivalent square"
    TestAssertClose stats, "model.squareEdgeAverage.target.rotation", _
        model.ConcreteRotation(2), edgeAngle, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Если матрица собственных инерций Region физически непригодна, габариты
' устойчивости определяются по квадрату той же площади.
Private Sub TestAutoCADImporterInvalidInertiaSquareFallback(ByRef stats As TTestStats)
    Dim concreteRegions(1 To 1, 1 To 7) As Variant
    concreteRegions(1, 1) = GEOM_PI * 25# * 25#
    concreteRegions(1, 2) = 0#
    concreteRegions(1, 3) = 0#
    concreteRegions(1, 4) = 100#
    concreteRegions(1, 5) = 100#
    concreteRegions(1, 6) = 10000#
    concreteRegions(1, 7) = "BAD_IXY"

    Dim rebarRegions(1 To 1, 1 To 4) As Variant
    rebarRegions(1, 1) = GEOM_PI * 12# * 12# / 4#
    rebarRegions(1, 2) = 0#
    rebarRegions(1, 3) = 0#
    rebarRegions(1, 4) = "RB"

    Dim importer As CAutoCADSectionModelImporter
    Set importer = New CAutoCADSectionModelImporter
    Dim model As CSectionModel
    Set model = importer.TestBuildFromRegionArrays(concreteRegions, rebarRegions, "Rebar", 0.000001)

    TestAssertTrue stats, "autocad.import.badInertia.shape", model.ConcreteShapeType(1) = "Equivalent square"
    TestAssertClose stats, "autocad.import.badInertia.localIxy", model.ConcreteLocalIxy(1), 0#, 0.000001

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    Dim minProjection As Double
    Dim maxProjection As Double
    props.CalculateProjection model, 1#, 0#, False, minProjection, maxProjection
    TestAssertClose stats, "autocad.import.badInertia.xDepth", maxProjection - minProjection, Sqr(model.ConcreteArea(1)), 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет произвольный неизотропный Region: оболочка сохраняет площадь и
' отношение главных I, но реальные локальные I остаются расчетными данными.
Private Sub TestSectionModelEquivalentRectangleFromRealInertia(ByRef stats As TTestStats)
    Dim area As Double
    Dim i1 As Double
    Dim i2 As Double
    Dim angle As Double
    Dim localIx As Double
    Dim localIy As Double
    Dim localIxy As Double
    area = 2500#
    i1 = 900000#
    i2 = 150000#
    angle = GEOM_PI / 8#
    TestRotatedLocalInertia i2, i1, angle, localIx, localIy, localIxy

    Dim model As CSectionModel
    Set model = New CSectionModel
    model.AddConcreteElement 0#, 0#, area, 1, "arbitrary", "H1", "Region", _
        0#, 0#, 0#, vbNullString, localIx, localIy, localIxy

    TestAssertTrue stats, "model.boundary.equivalentRectangle.status", _
        model.ConcreteShapeType(1) = "Equivalent rectangle"
    TestAssertClose stats, "model.boundary.equivalentRectangle.area", _
        model.ConcreteWidth(1) * model.ConcreteHeight(1), area, 0.000001
    TestAssertClose stats, "model.boundary.equivalentRectangle.ratio", _
        (model.ConcreteHeight(1) / model.ConcreteWidth(1)) ^ 2, i1 / i2, 0.000001
    TestAssertClose stats, "model.boundary.equivalentRectangle.realIx", model.ConcreteLocalIx(1), localIx, 0.000001
    TestAssertClose stats, "model.boundary.equivalentRectangle.realIy", model.ConcreteLocalIy(1), localIy, 0.000001
    TestAssertClose stats, "model.boundary.equivalentRectangle.realIxy", model.ConcreteLocalIxy(1), localIxy, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Изображение элемента не должно заменять реальные местные моменты инерции,
' по которым вычисляются свойства бетонного сечения.
Private Sub TestSectionModelBoundaryDoesNotReplaceRealInertia(ByRef stats As TTestStats)
    Dim area As Double
    Dim realIx As Double
    Dim realIy As Double
    Dim realIxy As Double
    area = 1000#
    realIx = 111111#
    realIy = 222222#
    realIxy = 12345#

    Dim modelA As CSectionModel
    Set modelA = New CSectionModel
    modelA.AddConcreteElement 0#, 0#, area, 1, "A", "A", "Rectangle", _
        100#, 10#, 0#, vbNullString, realIx, realIy, realIxy

    Dim modelB As CSectionModel
    Set modelB = New CSectionModel
    modelB.AddConcreteElement 0#, 0#, area, 1, "B", "B", "Rectangle", _
        10#, 100#, GEOM_PI / 3#, vbNullString, realIx, realIy, realIxy

    Dim propsA As CSectionPropertiesCalculator
    Set propsA = New CSectionPropertiesCalculator
    propsA.CalculateConcrete modelA

    Dim propsB As CSectionPropertiesCalculator
    Set propsB = New CSectionPropertiesCalculator
    propsB.CalculateConcrete modelB

    TestAssertClose stats, "model.boundary.realIx.priorityA", propsA.Ixc, realIx, 0.000001
    TestAssertClose stats, "model.boundary.realIy.priorityA", propsA.Iyc, realIy, 0.000001
    TestAssertClose stats, "model.boundary.realIxy.priorityA", propsA.Ixyc, realIxy, 0.000001
    TestAssertClose stats, "model.boundary.realIx.independent", propsB.Ixc, propsA.Ixc, 0.000001
    TestAssertClose stats, "model.boundary.realIy.independent", propsB.Iyc, propsA.Iyc, 0.000001
    TestAssertClose stats, "model.boundary.realIxy.independent", propsB.Ixyc, propsA.Ixyc, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' У элемента без сохранённых инерций CSectionModel использует квадрат
' той же площади и возвращает физически корректные моменты инерции.
Private Sub TestSectionModelFallbackInertiaSquare(ByRef stats As TTestStats)
    Dim area As Double
    area = 3600#

    Dim model As CSectionModel
    Set model = New CSectionModel
    model.AddConcreteElement 0#, 0#, area, 1, "noI", "noI", "Region"

    TestAssertTrue stats, "model.fallback.square.status", model.ConcreteShapeType(1) = "Equivalent square"
    TestAssertClose stats, "model.fallback.square.width", model.ConcreteWidth(1), 60#, 0.000001
    TestAssertClose stats, "model.fallback.square.localIx", model.ConcreteLocalIx(1), area * area / 12#, 0.000001
    TestAssertClose stats, "model.fallback.square.localIy", model.ConcreteLocalIy(1), area * area / 12#, 0.000001
    TestAssertClose stats, "model.fallback.square.localIxy", model.ConcreteLocalIxy(1), 0#, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает сеточный прямоугольник с аналитическими A, центром и Ix/Iy/Ixy.
' Скругления равны нулю, поэтому ошибка проверяет дискретизацию и перенос
' характеристик, а не приближение криволинейного контура.
Private Sub TestRectangle(ByRef stats As TTestStats)
    Dim geom As CGeometryRegion
    Set geom = TestFixtureBox(200#, 100#)

    Dim props As CSectionPropertiesCalculator
    Set props = TestMeshProps(geom, 2.5, 2.5)

    TestAssertClose stats, "rect.area", props.Area, 200# * 100#, 0.000001
    TestAssertClose stats, "rect.cx", props.CentroidX, 0#, 0.000001
    TestAssertClose stats, "rect.cy", props.CentroidY, 0#, 0.000001
    TestAssertRelative stats, "rect.Ix", props.Ixc, 200# * 100# ^ 3 / 12#, 0.001
    TestAssertRelative stats, "rect.Iy", props.Iyc, 100# * 200# ^ 3 / 12#, 0.001
    TestAssertClose stats, "rect.Ixy", props.Ixyc, 0#, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет второй/третий ряды Rectangles в Stack/SideBySide, разные диаметры
' двух линий, отсутствие первого ряда и частоту EverySecondBar.
' Координаты защищают направление смещения и перескок через второй ряд.
Private Sub TestRectanglesAdditionalRebarRows(ByRef stats As TTestStats)
    Dim builder As CRebarRuleExecutor
    Set builder = New CRebarRuleExecutor

    Dim stacked As CRebarLayout
    Set stacked = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(Array(50#, 50#, 20#, 20#, 1, 0, 100#, 100#, 100#, 100#, 20#, 0#, 30#, 0#, "Stacked", "Stacked"), TestEmptyFaceSettings(), TestEmptyFaceSettings(), TestEmptyFaceSettings())), builder)

    TestAssertTrue stats, "rectangles.rows.stacked.count", stacked.Count = 3
    TestAssertClose stats, "rectangles.rows.stacked.row1.x", stacked.X(1), 50#, 0.000001
    TestAssertClose stats, "rectangles.rows.stacked.row2.x", stacked.X(2), 50# + 20# + 2.5, 0.000001
    TestAssertClose stats, "rectangles.rows.stacked.row3.x", stacked.X(3), 50# + 20# + 2.5 + 25# + 2.5, 0.000001
    TestAssertClose stats, "rectangles.rows.stacked.sameY", stacked.Y(3), stacked.Y(1), 0.000001

    Dim sideVertical As CRebarLayout
    Set sideVertical = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(Array(50#, 50#, 20#, 20#, 1, 0, 100#, 100#, 100#, 100#, 20#, 0#, 30#, 0#, "SideBySide", "SideBySide"), TestEmptyFaceSettings(), TestEmptyFaceSettings(), TestEmptyFaceSettings())), builder)

    TestAssertTrue stats, "rectangles.rows.side.vertical.count", sideVertical.Count = 3
    TestAssertClose stats, "rectangles.rows.side.vertical.sameX", sideVertical.X(3), sideVertical.X(1), 0.000001
    TestAssertClose stats, "rectangles.rows.side.vertical.row2.y", sideVertical.Y(2), sideVertical.Y(1) - (20# + 2.5), 0.000001
    TestAssertClose stats, "rectangles.rows.side.vertical.row3.y", sideVertical.Y(3), sideVertical.Y(1) - (20# + 2.5 + 25# + 2.5), 0.000001

    Dim sideHorizontal As CRebarLayout
    Set sideHorizontal = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(TestEmptyFaceSettings(), TestEmptyFaceSettings(), Array(50#, 50#, 20#, 20#, 1, 0, 50#, 50#, 50#, 50#, 20#, 0#, 30#, 0#, "SideBySide", "SideBySide"), TestEmptyFaceSettings())), builder)

    TestAssertTrue stats, "rectangles.rows.side.horizontal.count", sideHorizontal.Count = 3
    TestAssertClose stats, "rectangles.rows.side.horizontal.row2.x", sideHorizontal.X(2), sideHorizontal.X(1) + 20# + 2.5, 0.000001
    TestAssertClose stats, "rectangles.rows.side.horizontal.row3.x", sideHorizontal.X(3), sideHorizontal.X(1) + 20# + 2.5 + 25# + 2.5, 0.000001
    TestAssertClose stats, "rectangles.rows.side.horizontal.sameY", sideHorizontal.Y(3), sideHorizontal.Y(1), 0.000001

    Dim thirdOnly As CRebarLayout
    Set thirdOnly = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(Array(50#, 50#, 20#, 20#, 1, 0, 100#, 100#, 100#, 100#, 0#, 0#, 30#, 0#, "Stacked", "Stacked"), TestEmptyFaceSettings(), TestEmptyFaceSettings(), TestEmptyFaceSettings())), builder)

    TestAssertTrue stats, "rectangles.rows.thirdOnly.count", thirdOnly.Count = 2
    TestAssertClose stats, "rectangles.rows.thirdOnly.row3.x", thirdOnly.X(2), thirdOnly.X(1) + 25# + 2.5, 0.000001
    TestAssertClose stats, "rectangles.rows.thirdOnly.sameY", thirdOnly.Y(2), thirdOnly.Y(1), 0.000001

    Dim noFirstRow As CRebarLayout
    Set noFirstRow = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(Array(50#, 50#, 20#, 20#, 0, 0, 100#, 100#, 100#, 100#, 20#, 20#, 30#, 30#, "Stacked", "Stacked"), Array(50#, 50#, 20#, 20#, 1, 0, 100#, 100#, 100#, 100#, 0#, 0#, 0#, 0#, "Stacked", "Stacked"), TestEmptyFaceSettings(), TestEmptyFaceSettings())), builder)

    TestAssertTrue stats, "rectangles.rows.noFirstRow.ignored", noFirstRow.Count = 1
    TestAssertTrue stats, "rectangles.rows.noFirstRow.source", InStr(1, noFirstRow.BarID(1), "H2", vbTextCompare) > 0

    Dim mixed As CRebarLayout
    Set mixed = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(Array(50#, 50#, 20#, 20#, 1, 0, 100#, 100#, 100#, 100#, 20#, 0#, 30#, 0#, "SideBySide", "Stacked"), TestEmptyFaceSettings(), TestEmptyFaceSettings(), TestEmptyFaceSettings())), builder)

    TestAssertTrue stats, "rectangles.rows.mixed.count", mixed.Count = 3
    TestAssertClose stats, "rectangles.rows.mixed.row2.y", mixed.Y(2), mixed.Y(1) - (20# + 2.5), 0.000001
    TestAssertClose stats, "rectangles.rows.mixed.row3.x", mixed.X(3), mixed.X(1) + 25# + 2.5, 0.000001
    TestAssertClose stats, "rectangles.rows.mixed.row3.y", mixed.Y(3), mixed.Y(1), 0.000001

    Dim separateLineDiameters As CRebarLayout
    Set separateLineDiameters = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(Array(50#, 50#, 20#, 20#, 1, 1, 100#, 100#, 100#, 100#, 18#, 22#, 28#, 34#, "Stacked", "Stacked"), TestEmptyFaceSettings(), TestEmptyFaceSettings(), TestEmptyFaceSettings())), builder)

    TestAssertTrue stats, "rectangles.rows.separateLineDiameters.count", separateLineDiameters.Count = 6
    TestAssertClose stats, "rectangles.rows.separateLineDiameters.row2_1.d", separateLineDiameters.Diameter(2), 18#, 0.000001
    TestAssertClose stats, "rectangles.rows.separateLineDiameters.row3_1.d", separateLineDiameters.Diameter(3), 28#, 0.000001
    TestAssertClose stats, "rectangles.rows.separateLineDiameters.row2_2.d", separateLineDiameters.Diameter(5), 22#, 0.000001
    TestAssertClose stats, "rectangles.rows.separateLineDiameters.row3_2.d", separateLineDiameters.Diameter(6), 34#, 0.000001

    Dim everySecond As CRebarLayout
    Set everySecond = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(Array(50#, 50#, 20#, 20#, 5, 0, 50#, 50#, 50#, 50#, 20#, 0#, 30#, 0#, "Stacked", "Stacked", "EverySecondBar", "EachBar"), TestEmptyFaceSettings(), TestEmptyFaceSettings(), TestEmptyFaceSettings())), builder)

    TestAssertTrue stats, "rectangles.rows.everySecond.count", everySecond.Count = 13
    TestAssertTrue stats, "rectangles.rows.everySecond.row2.partial", TestCountBarsWithRow(everySecond, ".Row2.") = 3
    TestAssertTrue stats, "rectangles.rows.everySecond.row3.each", TestCountBarsWithRow(everySecond, ".Row3.") = 5

    TestAssertRectanglesRebarRowsError stats, "rectangles.rows.invalid.location", _
        Array(50#, 50#, 20#, 20#, 1, 0, 100#, 100#, 100#, 100#, 20#, 0#, 0#, 0#, "Diagonal", "Stacked")
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет аналитические площадь/центр Г-сечения и точки его реального контура.
' На грубой сетке сравнивает центральную и subcell-дискретизацию, проверяя
' малые краевые элементы и сохранение расчетной площади.
Private Sub TestRectanglesGeometry(ByRef stats As TTestStats)
    Dim geom As CGeometryRegion
    Set geom = TestFixtureBoxes(Array(Array(10#, -20#, 600#, 250#), Array(10#, 230#, 250#, 550#)))

    Dim available As Boolean
    Dim analyticalArea As Double
    Dim centroidX As Double
    Dim centroidY As Double
    analyticalArea = geom.AnalyticalArea(available)
    geom.AnalyticalCentroid available, centroidX, centroidY

    TestAssertTrue stats, "rectangles.area.available", available
    TestAssertClose stats, "rectangles.area.analytical", analyticalArea, 600# * 250# + 250# * 550#, 0.000001
    TestAssertClose stats, "rectangles.cx.analytical", centroidX, 226.304347826087, 0.000001
    TestAssertClose stats, "rectangles.cy.analytical", centroidY, 296.304347826087, 0.000001
    TestAssertTrue stats, "rectangles.contains.lower", geom.ContainsPoint(580#, 20#)
    TestAssertTrue stats, "rectangles.contains.vertical", geom.ContainsPoint(100#, 700#)
    TestAssertTrue stats, "rectangles.excludes.cutout", Not geom.ContainsPoint(500#, 700#)

    Dim centerMesh As CFiberMeshBuilder
    Set centerMesh = New CFiberMeshBuilder
    centerMesh.BuildMesh geom, 80#, 80#, 1

    Dim subcellMesh As CFiberMeshBuilder
    Set subcellMesh = New CFiberMeshBuilder
    subcellMesh.BuildMesh geom, 80#, 80#, 1, 4

    Dim centerProps As CSectionPropertiesCalculator
    Set centerProps = New CSectionPropertiesCalculator
    centerProps.CalculateConcrete TestBuildGeneratedSectionModel(centerMesh, Nothing)

    Dim subcellProps As CSectionPropertiesCalculator
    Set subcellProps = New CSectionPropertiesCalculator
    subcellProps.CalculateConcrete TestBuildGeneratedSectionModel(subcellMesh, Nothing)

    TestAssertTrue stats, "rectangles.boundary.subcell.fibers.more", subcellMesh.FiberCount > centerMesh.FiberCount
    TestAssertTrue stats, "rectangles.boundary.subcell.has.small.fibers", TestMeshHasSmallFibers(subcellMesh, 80#)
    TestAssertRelative stats, "rectangles.boundary.subcell.area", subcellProps.Area, analyticalArea, 0.04
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет новые режимы Rectangles без изменения старого LSection: один
' прямоугольник, два прямоугольника без смещения и два прямоугольника со
' смещением верхнего блока по X.
Private Sub TestRectanglesGeometryModes(ByRef stats As TTestStats)
    Dim rectangle As CGeometryRegion
    Set rectangle = TestFixtureGeometry(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "OneRect")).ConcreteDomain)

    Dim available As Boolean
    Dim cx As Double
    Dim cy As Double
    TestAssertClose stats, "rectangles.rectangle.area", rectangle.AnalyticalArea(available), 250# * 550#, 0.000001
    TestAssertTrue stats, "rectangles.rectangle.area.available", available
    rectangle.AnalyticalCentroid available, cx, cy
    TestAssertClose stats, "rectangles.rectangle.cx", cx, 125#, 0.000001
    TestAssertClose stats, "rectangles.rectangle.cy", cy, 275#, 0.000001
    TestAssertTrue stats, "rectangles.rectangle.hasNoLower", rectangle.MinY = 0# And rectangle.MaxY = 550#
    TestAssertTrue stats, "rectangles.rectangle.containsUpper", rectangle.ContainsPoint(100#, 500#)
    TestAssertTrue stats, "rectangles.rectangle.excludesOldLower", Not rectangle.ContainsPoint(500#, 20#)

    Dim lsection As CGeometryRegion
    Dim twoZero As CGeometryRegion
    Set lsection = TestFixtureGeometry(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect")).ConcreteDomain)
    Set twoZero = TestFixtureGeometry(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect")).ConcreteDomain)
    TestAssertClose stats, "rectangles.twoZero.areaEqualsL", twoZero.AnalyticalArea(available), lsection.AnalyticalArea(available), 0.000001
    TestAssertClose stats, "rectangles.twoZero.minX", twoZero.MinX, lsection.MinX, 0.000001
    TestAssertClose stats, "rectangles.twoZero.maxX", twoZero.MaxX, lsection.MaxX, 0.000001
    TestAssertTrue stats, "rectangles.twoZero.cutout", Not twoZero.ContainsPoint(500#, 700#)

    Dim tshape As CGeometryRegion
    Set tshape = TestFixtureGeometry(TestFixtureRectanglesDefinition(Array(600#, 200#, 250#, 250#, -175#, "TwoRect")).ConcreteDomain)
    Dim validMessage As String
    TestAssertTrue stats, "rectangles.tshape.valid", tshape.IsValid(validMessage)
    TestAssertClose stats, "rectangles.tshape.minX", tshape.MinX, -175#, 0.000001
    TestAssertClose stats, "rectangles.tshape.maxX", tshape.MaxX, 425#, 0.000001
    TestAssertTrue stats, "rectangles.tshape.containsFlange", tshape.ContainsPoint(-100#, 300#)
    TestAssertTrue stats, "rectangles.tshape.containsWeb", tshape.ContainsPoint(125#, 100#)

    Dim shifted As CGeometryRegion
    Set shifted = TestFixtureGeometry(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 150#, "TwoRect")).ConcreteDomain)
    shifted.AnalyticalCentroid available, cx, cy
    Dim upperMinX As Double, upperMaxX As Double
    TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 150#, "TwoRect")).Path("Upper").ProjectionBounds 1#, 0#, upperMinX, upperMaxX
    TestAssertClose stats, "rectangles.shifted.upperMinX", upperMinX, 150#, 0.000001
    TestAssertClose stats, "rectangles.shifted.cx", cx, 288.04347826087, 0.000001
    TestAssertClose stats, "rectangles.shifted.cy", cy, 316.304347826087, 0.000001
    TestAssertTrue stats, "rectangles.shifted.containsUpper", shifted.ContainsPoint(250#, 700#)
    TestAssertTrue stats, "rectangles.shifted.excludesOldUpper", Not shifted.ContainsPoint(50#, 700#)

    Dim builder As CRebarRuleExecutor
    Set builder = New CRebarRuleExecutor
    Dim shiftedBars As CRebarLayout
    Set shiftedBars = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 150#, "TwoRect"), _
        Array(Array(50#, 50#, 20#, 20#, 1, 0, 80#, 80#, 80#, 80#), TestEmptyFaceSettings(), TestEmptyFaceSettings(), TestEmptyFaceSettings())), builder)
    TestAssertClose stats, "rectangles.shifted.rebarX", shiftedBars.X(1), 200#, 0.000001

    Dim rectangleBars As CRebarLayout
    Set rectangleBars = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "OneRect"), _
        Array(TestEmptyFaceSettings(), Array(50#, 50#, 20#, 20#, 2, 2, 80#, 80#, 80#, 80#), Array(50#, 50#, 20#, 20#, 2, 0, 80#, 80#, 80#, 80#), Array(50#, 50#, 20#, 20#, 5, 5, 80#, 80#, 80#, 80#))), builder)
    TestAssertTrue stats, "rectangles.rectangle.rebarSkipsLower", TestCountBarsWithSource(rectangleBars, "H2") = 0 And TestCountBarsWithSource(rectangleBars, "B2") = 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что PrincipalAngle задает физическую главную ось 1, а ядровые
' расстояния главных плоскостей считаются через ту же нормаль, что и
' устойчивость. Тест защищает Г-сечение, где ошибка знака угла сразу дает
' заметный ненулевой I12 после поворота.
Private Sub TestRectanglesPrincipalAxesAndCoreDistances(ByRef stats As TTestStats)
    Dim geom As CGeometryRegion
    Set geom = TestFixtureGeometry(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect")).ConcreteDomain)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 25#, 25#, 1, 2

    Dim concreteSection As CSectionModel
    Set concreteSection = TestBuildGeneratedSectionModel(mesh, Nothing, "RectanglesPrincipalConcrete")

    Dim concreteProps As CSectionPropertiesCalculator
    Set concreteProps = New CSectionPropertiesCalculator
    concreteProps.CalculateConcrete concreteSection
    TestAssertPrincipalAxesConsistent stats, "rectangles.concrete", concreteSection, concreteProps

    Dim rebarBuilder As CRebarRuleExecutor
    Set rebarBuilder = New CRebarRuleExecutor
    Dim rebars As CRebarLayout
    Set rebars = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(Array(40#, 40#, 32#, 32#, 5, 5, 80#, 80#, 80#, 80#), Array(40#, 40#, 32#, 32#, 2, 2, 80#, 80#, 80#, 80#), Array(40#, 40#, 32#, 32#, 2, 2, 80#, 80#, 80#, 80#), Array(40#, 40#, 32#, 32#, 5, 5, 80#, 80#, 80#, 80#))), rebarBuilder)

    Dim transformedSection As CSectionModel
    Set transformedSection = TestBuildGeneratedSectionModel(mesh, rebars, "RectanglesPrincipalTransformed")

    Dim transformedProps As CSectionPropertiesCalculator
    Set transformedProps = New CSectionPropertiesCalculator
    transformedProps.CalculateTransformedByModuli transformedSection, 32500#, 200000#
    TestAssertPrincipalAxesConsistent stats, "rectangles.transformed", transformedSection, transformedProps
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет общие инварианты главных осей для любого представления сечения.
' Ось 1 должна давать I1, ось 2 - I2, а I12 в повернутой системе должен
' исчезать. CoreDistance главной плоскости сверяется с прямым расчетом по
' соответствующей нормали к оси изгиба.
Private Sub TestAssertPrincipalAxesConsistent(ByRef stats As TTestStats, ByVal prefix As String, _
        ByVal section As CSectionModel, ByVal props As CSectionPropertiesCalculator)
    Dim axis1X As Double
    Dim axis1Y As Double
    Dim axis2X As Double
    Dim axis2Y As Double
    props.PrincipalAxisDirection 1, axis1X, axis1Y
    props.PrincipalAxisDirection 2, axis2X, axis2Y

    Dim inertia1 As Double
    Dim inertia2 As Double
    inertia1 = props.ConcreteInertiaAboutAxis(section, axis1X, axis1Y, props.CentroidX, props.CentroidY)
    inertia2 = props.ConcreteInertiaAboutAxis(section, axis2X, axis2Y, props.CentroidX, props.CentroidY)
    If section.RebarCount > 0 Then
        inertia1 = inertia1 + (200000# / 32500# - 1#) * props.RebarInertiaAboutAxis(section, axis1X, axis1Y, props.CentroidX, props.CentroidY)
        inertia2 = inertia2 + (200000# / 32500# - 1#) * props.RebarInertiaAboutAxis(section, axis2X, axis2Y, props.CentroidX, props.CentroidY)
    End If

    TestAssertRelative stats, prefix & ".axis1.inertia", inertia1, props.PrincipalI1, 0.0000001
    TestAssertRelative stats, prefix & ".axis2.inertia", inertia2, props.PrincipalI2, 0.0000001

    Dim c As Double
    Dim s As Double
    Dim rotatedIxy As Double
    c = Cos(props.PrincipalAngleRad)
    s = Sin(props.PrincipalAngleRad)
    rotatedIxy = (props.Ixc - props.Iyc) * s * c + props.Ixyc * (c * c - s * s)
    TestAssertTrue stats, prefix & ".principal.I12.zero", _
        Abs(rotatedIxy) <= GeomMax(props.PrincipalI1, 1#) * 0.0000001

    Dim normal1X As Double
    Dim normal1Y As Double
    props.PrincipalPlaneNormalDirection 1, normal1X, normal1Y
    TestAssertClose stats, prefix & ".core.plane1.plus", _
        props.PrincipalPlaneCoreDistance(section, 1, True, False), _
        props.CoreDistanceAlong(section, normal1X, normal1Y, False), 0.000001
    TestAssertClose stats, prefix & ".core.plane1.minus", _
        props.PrincipalPlaneCoreDistance(section, 1, False, False), _
        props.CoreDistanceAlong(section, -normal1X, -normal1Y, False), 0.000001

    Dim normal2X As Double
    Dim normal2Y As Double
    props.PrincipalPlaneNormalDirection 2, normal2X, normal2Y
    TestAssertClose stats, prefix & ".core.plane2.plus", _
        props.PrincipalPlaneCoreDistance(section, 2, True, False), _
        props.CoreDistanceAlong(section, normal2X, normal2Y, False), 0.000001
    TestAssertClose stats, prefix & ".core.plane2.minus", _
        props.PrincipalPlaneCoreDistance(section, 2, False, False), _
        props.CoreDistanceAlong(section, -normal2X, -normal2Y, False), 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет локальный поиск бетонной границы для a_s. На ступенчатом контуре
' глобальная опорная линия всего сечения лежит на верхнем выступе, но луч из
' точки должен выйти через ближайшую грань той ветви, где находится стержень.
Private Sub TestConcreteCoverUsesLocalContour(ByRef stats As TTestStats)
    Dim section As CSectionModel
    Set section = New CSectionModel
    section.AddConcreteElement 0#, 0#, 80000#, 1, "", "", "Rectangle", 400#, 200#, 0#
    section.AddConcreteElement -150#, 200#, 40000#, 1, "", "", "Rectangle", 100#, 400#, 0#

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator

    Dim query As CSectionGeometryQuery
    Set query = New CSectionGeometryQuery
    query.Initialize section

    Dim localCover As Double
    localCover = query.ConcreteCoverFromPointAlongDirection(0#, 60#, 0#, 1#)
    TestAssertClose stats, "geometry.cover.localContour", localCover, 40#, 0.000001

    Dim minS As Double
    Dim maxS As Double
    props.CalculateProjection section, 0#, 1#, False, minS, maxS
    TestAssertClose stats, "geometry.cover.globalWouldBeWrong", maxS - 60#, 340#, 0.000001
    TestAssertTrue stats, "geometry.cover.localLessThanGlobal", localCover < maxS - 60#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет количество, шаг и координаты автоматических гранных стержней.
' Нулевые число/диаметр исключают линию, остальные стержни остаются в бетоне;
' отрицательное число и невозможные отступы дают ошибку ввода.
Private Sub TestRectanglesAutoRebarLayout(ByRef stats As TTestStats)
    Dim builder As CRebarRuleExecutor
    Set builder = New CRebarRuleExecutor

    Dim layout As CRebarLayout
    Set layout = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(Array(50#, 50#, 20#, 20#, 3, 0, 100#, 100#, 100#, 100#), Array(50#, 50#, 20#, 20#, 0, 0, 50#, 50#, 50#, 50#), Array(50#, 50#, 20#, 20#, 2, 0, 60#, 60#, 60#, 60#), Array(50#, 50#, 20#, 20#, 3, 0, 100#, 100#, 100#, 100#))), builder)

    TestAssertTrue stats, "rectangles.rebar.count", layout.Count = 8
    TestAssertTrue stats, "rectangles.rebar.requested", builder.RequestedPositions = 8
    Dim footprint As CSectionDefinition
    Set footprint = TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#))
    TestAssertClose stats, "rectangles.rebar.perimeter", footprint.Path("Upper").Length + footprint.Path("Lower").Length, 3300#, 0.000001
    TestAssertClose stats, "rectangles.rebar.outerPerimeter", footprint.Path("Boundary").Length, 2800#, 0.000001
    TestAssertClose stats, "rectangles.rebar.step", Abs(layout.X(7) - layout.X(6)), 200#, 0.000001
    TestAssertClose stats, "rectangles.rebar.firstX", layout.X(1), 50#, 0.000001
    TestAssertClose stats, "rectangles.rebar.firstY", layout.Y(1), 525#, 0.000001
    TestAssertTrue stats, "rectangles.rebar.zeroFaceSkipped", InStr(1, layout.BarID(1), "H2", vbTextCompare) = 0

    Dim noLineByCount As CRebarLayout
    Set noLineByCount = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(Array(50#, 50#, 0#, 20#, 0, 0, 100#, 100#, 100#, 100#), Array(50#, 50#, 20#, 20#, 0, 0, 50#, 50#, 50#, 50#), Array(50#, 50#, 20#, 20#, 2, 0, 60#, 60#, 60#, 60#), Array(50#, 50#, 20#, 20#, 3, 0, 100#, 100#, 100#, 100#))), builder)
    TestAssertTrue stats, "rectangles.rebar.zeroCountSkipped", noLineByCount.Count = 5

    Dim noLineByDiameter As CRebarLayout
    Set noLineByDiameter = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(Array(50#, 50#, 0#, 20#, 3, 0, 100#, 100#, 100#, 100#), Array(50#, 50#, 20#, 20#, 0, 0, 50#, 50#, 50#, 50#), Array(50#, 50#, 20#, 20#, 2, 0, 60#, 60#, 60#, 60#), Array(50#, 50#, 20#, 20#, 3, 0, 100#, 100#, 100#, 100#))), builder)
    TestAssertTrue stats, "rectangles.rebar.zeroDiameterSkipped", noLineByDiameter.Count = 5

    Dim noBarsByCount As CRebarLayout
    Set noBarsByCount = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(Array(50#, 50#, 0#, 0#, 0, 0, 100#, 100#, 100#, 100#), Array(50#, 50#, 0#, 0#, 0, 0, 50#, 50#, 50#, 50#), Array(50#, 50#, 0#, 0#, 0, 0, 60#, 60#, 60#, 60#), Array(50#, 50#, 0#, 0#, 0, 0, 100#, 100#, 100#, 100#))), builder)
    TestAssertTrue stats, "rectangles.rebar.allZeroCounts.builderEmpty", noBarsByCount.Count = 0

    Dim geom As CGeometryRegion
    Set geom = TestFixtureGeometry(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect")).ConcreteDomain)

    Dim i As Long
    For i = 1 To layout.Count
        TestAssertTrue stats, "rectangles.rebar.inside." & CStr(i), geom.ContainsPoint(layout.X(i), layout.Y(i))
    Next i

    TestAssertRectanglesRebarError stats, "rectangles.rebar.invalid.n", 600#, 550#, 250#, 250#, -1, 50#, 20#
    TestAssertRectanglesRebarError stats, "rectangles.rebar.invalid.as.small", 600#, 550#, 250#, 250#, 2, 10#, 20#
    TestAssertRectanglesRebarError stats, "rectangles.rebar.invalid.offsets", 600#, 550#, 250#, 250#, 2, 50#, 20#, 500#, 500#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задает разные начальные/конечные отступы противоположным линиям H1.
' Проверяет четыре координаты стержней, чтобы настройки одной линии
' не подменяли независимые настройки соседней линии.
Private Sub TestRectanglesSeparateLineOffsets(ByRef stats As TTestStats)
    Dim builder As CRebarRuleExecutor
    Set builder = New CRebarRuleExecutor

    Dim layout As CRebarLayout
    Set layout = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(Array(40#, 60#, 20#, 20#, 2, 2, 80#, 120#, 30#, 70#), Array(40#, 40#, 20#, 20#, 0, 0, 40#, 40#, 40#, 40#), Array(40#, 40#, 20#, 20#, 0, 0, 40#, 40#, 40#, 40#), Array(40#, 40#, 20#, 20#, 0, 0, 40#, 40#, 40#, 40#))), builder)

    TestAssertTrue stats, "rectangles.offsets.separate.count", layout.Count = 4
    TestAssertClose stats, "rectangles.offsets.as_1.firstY", layout.Y(1), 680#, 0.000001
    TestAssertClose stats, "rectangles.offsets.as_1.edgeY", layout.Y(2), 330#, 0.000001
    TestAssertClose stats, "rectangles.offsets.as_2.firstY", layout.Y(3), 320#, 0.000001
    TestAssertClose stats, "rectangles.offsets.as_2.edgeY", layout.Y(4), 770#, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет круговой шаг, радиус осей, центр, площадь и дополнительные ряды.
' Защищает направление Stack/SideBySide и пропуск зависимых рядов без первого;
' невозможные диаметр, защитное расстояние и расположение отклоняются.
Private Sub TestCircleAutoRebarLayout(ByRef stats As TTestStats)
    ' Начало окружности — на 12 часах; касательная направлена вправо.
    ' Контактные расстояния равны сумме радиусов и прибавке 2,5 мм.
    Dim builder As CRebarRuleExecutor
    Set builder = New CRebarRuleExecutor

    Dim layout As CRebarLayout
    Set layout = TestFixtureMoveLayout(TestFixtureAnnotatedLayout(TestFixtureCircleDefinition(Array(300#, 40#, 8, 20#)), builder), 10#, -20#)

    TestAssertTrue stats, "circle.rebar.count", layout.Count = 8
    TestAssertClose stats, "circle.rebar.angleStep", Atn((layout.Y(2) + 20#) / (layout.X(2) - 10#)), 2# * GEOM_PI / 8#, 0.000000001
    TestAssertTrue stats, "circle.rebar.clockwise.firstQuadrant", layout.X(2) > 10# And layout.Y(2) < layout.Y(1)
    TestAssertClose stats, "circle.rebar.axisRadius", Sqr((layout.X(1) - 10#) ^ 2 + (layout.Y(1) + 20#) ^ 2), 110#, 0.000001
    TestAssertClose stats, "circle.rebar.firstX", layout.X(1), 10#, 0.000001
    TestAssertClose stats, "circle.rebar.firstY", layout.Y(1), -20# + 110#, 0.000001
    TestAssertClose stats, "circle.rebar.centroidX", (layout.X(1) + layout.X(5)) / 2#, 10#, 0.000001
    TestAssertClose stats, "circle.rebar.centroidY", (layout.Y(1) + layout.Y(5)) / 2#, -20#, 0.000001
    TestAssertRelative stats, "circle.rebar.area", layout.Area(1), GEOM_PI * 20# * 20# / 4#, 0.000000001

    Dim stacked As CRebarLayout
    Set stacked = TestFixtureMoveLayout(TestFixtureAnnotatedLayout(TestFixtureCircleDefinition(Array(300#, 40#, 8, 20#, 20#, 30#, "Stacked", "Stacked")), builder), 10#, -20#)
    TestAssertTrue stats, "circle.rebar.rows.stacked.count", stacked.Count = 24
    TestAssertClose stats, "circle.rebar.rows.stacked.row2.x", stacked.X(2), 10#, 0.000001
    TestAssertClose stats, "circle.rebar.rows.stacked.row2.y", stacked.Y(2), 90# - ((20# + 20#) / 2# + 2.5), 0.000001
    TestAssertClose stats, "circle.rebar.rows.stacked.row3.x", stacked.X(3), 10#, 0.000001
    TestAssertClose stats, "circle.rebar.rows.stacked.row3.y", stacked.Y(3), 90# - ((20# + 20#) / 2# + 2.5) - ((20# + 30#) / 2# + 2.5), 0.000001

    Dim sideBySide As CRebarLayout
    Set sideBySide = TestFixtureMoveLayout(TestFixtureAnnotatedLayout(TestFixtureCircleDefinition(Array(300#, 40#, 8, 20#, 20#, 30#, "SideBySide", "SideBySide")), builder), 10#, -20#)
    TestAssertTrue stats, "circle.rebar.rows.side.count", sideBySide.Count = 24
    TestAssertClose stats, "circle.rebar.rows.side.row2.x", sideBySide.X(2), 10# + ((20# + 20#) / 2# + 2.5), 0.000001
    TestAssertClose stats, "circle.rebar.rows.side.row2.y", sideBySide.Y(2), 90#, 0.000001
    TestAssertClose stats, "circle.rebar.rows.side.row3.x", sideBySide.X(3), 10# + ((20# + 20#) / 2# + 2.5) + ((20# + 30#) / 2# + 2.5), 0.000001
    TestAssertClose stats, "circle.rebar.rows.side.row3.y", sideBySide.Y(3), 90#, 0.000001

    Dim mixed As CRebarLayout
    Set mixed = TestFixtureMoveLayout(TestFixtureAnnotatedLayout(TestFixtureCircleDefinition(Array(300#, 40#, 8, 20#, 20#, 30#, "SideBySide", "Stacked")), builder), 10#, -20#)
    TestAssertTrue stats, "circle.rebar.rows.mixed.count", mixed.Count = 24
    TestAssertClose stats, "circle.rebar.rows.mixed.row2.y", mixed.Y(2), 90#, 0.000001
    TestAssertClose stats, "circle.rebar.rows.mixed.row3.x", mixed.X(3), 10#, 0.000001
    TestAssertClose stats, "circle.rebar.rows.mixed.row3.y", mixed.Y(3), 90# - ((20# + 30#) / 2# + 2.5), 0.000001

    Dim thirdOnly As CRebarLayout
    Set thirdOnly = TestFixtureMoveLayout(TestFixtureAnnotatedLayout(TestFixtureCircleDefinition(Array(300#, 40#, 8, 20#, 0#, 30#, "Stacked", "Stacked")), builder), 10#, -20#)
    TestAssertTrue stats, "circle.rebar.rows.thirdOnly.count", thirdOnly.Count = 16
    TestAssertClose stats, "circle.rebar.rows.thirdOnly.row3.x", thirdOnly.X(2), 10#, 0.000001
    TestAssertClose stats, "circle.rebar.rows.thirdOnly.row3.y", thirdOnly.Y(2), 90# - ((20# + 30#) / 2# + 2.5), 0.000001

    Dim noFirstByCount As CRebarLayout
    Set noFirstByCount = TestFixtureMoveLayout(TestFixtureAnnotatedLayout(TestFixtureCircleDefinition(Array(300#, 40#, 0, 20#, 20#, 30#, "Stacked", "Stacked")), builder), 10#, -20#)
    TestAssertTrue stats, "circle.rebar.rows.noFirstByCount.ignored", noFirstByCount.Count = 0
    TestAssertClose stats, "circle.rebar.rows.noFirstByCount.angle", CDbl(builder.CreatedPositions), 0#, 0.000000001

    Dim noFirstByDiameter As CRebarLayout
    Set noFirstByDiameter = TestFixtureMoveLayout(TestFixtureAnnotatedLayout(TestFixtureCircleDefinition(Array(300#, 40#, 8, 0#, 20#, 30#, "SideBySide", "SideBySide")), builder), 10#, -20#)
    TestAssertTrue stats, "circle.rebar.rows.noFirstByDiameter.ignored", noFirstByDiameter.Count = 0

    TestAssertCircleRebarError stats, "circle.rebar.invalid.D", 0#, 40#, 8, 20#
    TestAssertCircleRebarError stats, "circle.rebar.invalid.n", 300#, 40#, 1, 20#
    TestAssertCircleRebarError stats, "circle.rebar.invalid.ds", 300#, 40#, 8, -1#
    TestAssertCircleRebarError stats, "circle.rebar.invalid.as.small", 300#, 10#, 8, 20#
    TestAssertCircleRebarError stats, "circle.rebar.invalid.as.large", 300#, 150#, 8, 20#
    TestAssertCircleRebarError stats, "circle.rebar.invalid.loc2", 300#, 40#, 8, 20#, 20#, 0#, "Diagonal", "Stacked"
    TestAssertCircleRebarError stats, "circle.rebar.invalid.row2.outside", 300#, 20#, 8, 20#, 40#, 0#, "SideBySide", "Stacked"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает две сетки круга: попадание центров и разбиение краевых ячеек.
' Subcells должны улучшить площадь, увеличить число граничных волокон
' и сохранить их настоящий уменьшенный размер в модели.
Private Sub TestBoundarySubcellMesh(ByRef stats As TTestStats)
    Dim geom As CGeometryRegion
    Set geom = TestFixtureDisk(190#)

    Dim centerMesh As CFiberMeshBuilder
    Set centerMesh = New CFiberMeshBuilder
    centerMesh.BuildMesh geom, 40#, 40#, 1

    Dim subcellMesh As CFiberMeshBuilder
    Set subcellMesh = New CFiberMeshBuilder
    subcellMesh.BuildMesh geom, 40#, 40#, 1, 4

    Dim centerProps As CSectionPropertiesCalculator
    Set centerProps = New CSectionPropertiesCalculator
    centerProps.CalculateConcrete TestBuildGeneratedSectionModel(centerMesh, Nothing)

    Dim subcellProps As CSectionPropertiesCalculator
    Set subcellProps = New CSectionPropertiesCalculator
    subcellProps.CalculateConcrete TestBuildGeneratedSectionModel(subcellMesh, Nothing)

    Dim targetArea As Double
    targetArea = GEOM_PI * 95# * 95#

    TestAssertTrue stats, "boundary.subcell.fibers.more", subcellMesh.FiberCount > centerMesh.FiberCount
    TestAssertTrue stats, "boundary.subcell.area.better", Abs(subcellProps.Area - targetArea) < Abs(centerProps.Area - targetArea)
    TestAssertTrue stats, "boundary.subcell.has.small.fibers", TestMeshHasSmallFibers(subcellMesh, 40#)
    TestAssertClose stats, "boundary.subcell.small.width", TestFirstSmallFiberWidth(subcellMesh, 40#), 10#, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что генератор бетонной сетки действительно поддерживает разные
' шаги по X и Y и передает прямоугольную оболочку элемента дальше в модель.
Private Sub TestRectangularMeshSteps(ByRef stats As TTestStats)
    Dim geom As CGeometryRegion
    Set geom = TestFixtureBox(100#, 60#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 25#, 15#, 1, 1

    TestAssertTrue stats, "mesh.rectangularSteps.hasFibers", mesh.FiberCount > 0
    TestAssertClose stats, "mesh.rectangularSteps.width", mesh.FiberWidth(1), 25#, 0.000001
    TestAssertClose stats, "mesh.rectangularSteps.height", mesh.FiberHeight(1), 15#, 0.000001
    TestAssertClose stats, "mesh.rectangularSteps.area", mesh.FiberArea(1), 25# * 15#, 0.000001

    Dim model As CSectionModel
    Set model = TestBuildGeneratedSectionModel(mesh, Nothing)
    TestAssertClose stats, "mesh.rectangularSteps.modelWidth", model.ConcreteWidth(1), 25#, 0.000001
    TestAssertClose stats, "mesh.rectangularSteps.modelHeight", model.ConcreteHeight(1), 15#, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет круг с ненулевым центром: аналитические размеры/характеристики,
' включение точек в контур и сеточные A/I относительно независимых формул.
' Координаты центра не должны теряться при генерации модели.
Private Sub TestCircleGeometry(ByRef stats As TTestStats)
    Dim geom As CGeometryRegion
    Set geom = TestFixtureDisk(300#, 10#, -20#)

    Dim available As Boolean
    Dim analyticalArea As Double
    Dim centroidX As Double
    Dim centroidY As Double
    analyticalArea = geom.AnalyticalArea(available)
    geom.AnalyticalCentroid available, centroidX, centroidY

    Dim props As CSectionPropertiesCalculator
    Set props = TestMeshProps(geom, 2.5, 2.5)

    Dim analyticalI As Double
    analyticalI = GEOM_PI * 150# ^ 4 / 4#

    TestAssertTrue stats, "circle.area.available", available
    TestAssertRelative stats, "circle.area", props.Area, analyticalArea, 0.003
    TestAssertClose stats, "circle.cx", props.CentroidX, centroidX, 0.05
    TestAssertClose stats, "circle.cy", props.CentroidY, centroidY, 0.05
    TestAssertRelative stats, "circle.Ix", props.Ixc, analyticalI, 0.006
    TestAssertRelative stats, "circle.Iy", props.Iyc, analyticalI, 0.006
    TestAssertClose stats, "circle.Ixy", props.Ixyc, 0#, 0.000001 * props.Area * 150#
    TestAssertClose stats, "circle.principal.angle", props.PrincipalAngleRad, 0#, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет ядровое расстояние на круге: для сплошного круга r = R/4.
' Метод CoreDistanceAlong затем используется устойчивостью СП 35 для ветви e/r.
Private Sub TestCircleCoreDistance(ByRef stats As TTestStats)
    Dim geom As CGeometryRegion
    Set geom = TestFixtureDisk(300#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 2.5, 2.5, 1

    Dim section As CSectionModel
    Set section = TestBuildGeneratedSectionModel(mesh, Nothing)

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete section

    TestAssertRelative stats, "circle.core.xPositive", props.CoreDistanceAlong(section, 1#, 0#, False), 37.5, 0.02
    TestAssertRelative stats, "circle.core.yNegative", props.CoreDistanceAlong(section, 0#, -1#, False), 37.5, 0.02
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет соглашение Plus/Minus на несимметричном сечении. Направление
' CoreDistanceAlong задает направление эксцентриситета, поэтому для +X в
' формуле r = I/(A*c) используется крайнее волокно с противоположной стороны,
' то есть расстояние от центра тяжести до левой границы.
Private Sub TestAsymmetricCoreDistanceUsesOppositeFiber(ByRef stats As TTestStats)
    Dim section As CSectionModel
    Set section = New CSectionModel
    section.AddConcreteElement 0#, 0#, 10000#, 1, "", "", "Rectangle", 100#, 100#, 0#
    section.AddConcreteElement 125#, 0#, 5000#, 1, "", "", "Rectangle", 50#, 100#, 0#

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete section

    Dim minProjection As Double
    Dim maxProjection As Double
    props.CalculateProjection section, 1#, 0#, False, minProjection, maxProjection

    Dim inertiaForX As Double
    Dim expectedXPlus As Double
    Dim expectedXMinus As Double
    inertiaForX = props.ProjectedInertiaAboutCentroid(1#, 0#)
    expectedXPlus = inertiaForX / (props.Area * (props.CentroidX - minProjection))
    expectedXMinus = inertiaForX / (props.Area * (maxProjection - props.CentroidX))

    TestAssertTrue stats, "core.asymmetric.expectedDifferent", Abs(expectedXPlus - expectedXMinus) > 0.000001
    TestAssertClose stats, "core.asymmetric.xPlus.oppositeFiber", _
        props.CoreDistanceAlong(section, 1#, 0#, False), expectedXPlus, 0.000001
    TestAssertClose stats, "core.asymmetric.xMinus.oppositeFiber", _
        props.CoreDistanceAlong(section, -1#, 0#, False), expectedXMinus, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет общий случай центральных осей, когда Ixy не равен нулю. Для
' несимметричного Г-сечения ядровое расстояние по глобальным X/Y нельзя
' получать как I/(A*c) без учета произведения инерции: граница ядра должна
' обнулять линейный множитель напряжений по полной матрице I.
Private Sub TestCoreDistanceWithProductInertiaMatchesLinearBoundary(ByRef stats As TTestStats)
    Dim geom As CGeometryRegion
    Set geom = TestFixtureGeometry(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect")).ConcreteDomain)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 25#, 25#, 1, 2

    Dim rebarBuilder As CRebarRuleExecutor
    Set rebarBuilder = New CRebarRuleExecutor

    Dim rebars As CRebarLayout
    Set rebars = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#, 0#, "TwoRect"), _
        Array(Array(40#, 40#, 32#, 32#, 5, 5, 80#, 80#, 80#, 80#), Array(40#, 40#, 32#, 32#, 2, 2, 80#, 80#, 80#, 80#), Array(40#, 40#, 32#, 32#, 2, 2, 80#, 80#, 80#, 80#), Array(40#, 40#, 32#, 32#, 5, 5, 80#, 80#, 80#, 80#))), rebarBuilder)

    Dim section As CSectionModel
    Set section = TestBuildGeneratedSectionModel(mesh, rebars, "RectanglesCoreDistance")

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateTransformedByModuli section, 32500#, 200000#

    TestAssertCoreDistanceBoundary stats, "core.general.xPlus", section, props, 1#, 0#
    TestAssertCoreDistanceBoundary stats, "core.general.xMinus", section, props, -1#, 0#
    TestAssertCoreDistanceBoundary stats, "core.general.yPlus", section, props, 0#, 1#
    TestAssertCoreDistanceBoundary stats, "core.general.yMinus", section, props, 0#, -1#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет r с прямым линейным условием ядра: при 0.95*r вся граница еще
' сжата, при r крайняя точка имеет нулевую деформацию, при 1.05*r появляется
' растяжение.
Private Sub TestAssertCoreDistanceBoundary(ByRef stats As TTestStats, ByVal prefix As String, _
        ByVal section As CSectionModel, ByVal props As CSectionPropertiesCalculator, _
        ByVal directionX As Double, ByVal directionY As Double)
    Dim rValue As Double
    rValue = props.CoreDistanceAlong(section, directionX, directionY, False)

    Dim insideValue As Double
    Dim edgeValue As Double
    Dim outsideValue As Double
    insideValue = TestLinearKernelMinimum(section, props, directionX, directionY, 0.95 * rValue)
    edgeValue = TestLinearKernelMinimum(section, props, directionX, directionY, rValue)
    outsideValue = TestLinearKernelMinimum(section, props, directionX, directionY, 1.05 * rValue)

    TestAssertTrue stats, prefix & ".inside", insideValue > 0#
    TestAssertClose stats, prefix & ".edge", edgeValue, 0#, 0.000000000001
    TestAssertTrue stats, prefix & ".outside", outsideValue < 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает минимальное по бетонной оболочке значение скобки:
' 1/A + e * g(x,y). Для сжатой N вся скобка должна оставаться
' неотрицательной; отрицательное значение означает появление растяжения.
Private Function TestLinearKernelMinimum(ByVal section As CSectionModel, _
        ByVal props As CSectionPropertiesCalculator, ByVal directionX As Double, _
        ByVal directionY As Double, ByVal eccentricity As Double) As Double
    Dim normValue As Double
    normValue = Sqr(directionX * directionX + directionY * directionY)

    Dim nx As Double
    Dim ny As Double
    nx = directionX / normValue
    ny = directionY / normValue

    Dim determinant As Double
    determinant = props.Ixc * props.Iyc - props.Ixyc * props.Ixyc

    Dim kxPerEccentricity As Double
    Dim kyPerEccentricity As Double
    kxPerEccentricity = (props.Iyc * ny - props.Ixyc * nx) / determinant
    kyPerEccentricity = (-props.Ixyc * ny + props.Ixc * nx) / determinant

    Dim minProjection As Double
    minProjection = TestMinimumLinearBoundaryProjection(section, kyPerEccentricity, kxPerEccentricity)

    TestLinearKernelMinimum = 1# / props.Area + eccentricity * _
        (minProjection - kyPerEccentricity * props.CentroidX - kxPerEccentricity * props.CentroidY)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ищет минимум линейного функционала axisX*x+axisY*y по расчетной бетонной
' оболочке. Для прямоугольной оболочки линейная функция достигает экстремума
' в одном из углов, поэтому достаточно перебрать углы каждого бетонного КЭ.
Private Function TestMinimumLinearBoundaryProjection(ByVal section As CSectionModel, _
        ByVal axisX As Double, ByVal axisY As Double) As Double
    Dim initialized As Boolean
    Dim minValue As Double
    Dim i As Long
    For i = 1 To section.ConcreteCount
        Dim widthValue As Double
        Dim heightValue As Double
        Dim angleValue As Double
        If section.ConcreteBoundaryRectangle(i, widthValue, heightValue, angleValue) Then
            TestIncludeRectangleProjectionMinimum section.ConcreteX(i), section.ConcreteY(i), _
                widthValue, heightValue, angleValue, axisX, axisY, initialized, minValue
        Else
            TestIncludeProjectionMinimum axisX * section.ConcreteX(i) + axisY * section.ConcreteY(i), _
                initialized, minValue
        End If
    Next i
    TestMinimumLinearBoundaryProjection = minValue
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Добавляет четыре угла бетонной оболочки в поиск минимума линейной формы.
Private Sub TestIncludeRectangleProjectionMinimum(ByVal centerX As Double, ByVal centerY As Double, _
        ByVal widthValue As Double, ByVal heightValue As Double, ByVal angleValue As Double, _
        ByVal axisX As Double, ByVal axisY As Double, ByRef initialized As Boolean, _
        ByRef minValue As Double)
    Dim ux As Double
    Dim uy As Double
    Dim vx As Double
    Dim vy As Double
    ux = Cos(angleValue)
    uy = Sin(angleValue)
    vx = -uy
    vy = ux

    TestIncludeProjectionMinimum axisX * (centerX + ux * widthValue / 2# + vx * heightValue / 2#) + _
        axisY * (centerY + uy * widthValue / 2# + vy * heightValue / 2#), initialized, minValue
    TestIncludeProjectionMinimum axisX * (centerX + ux * widthValue / 2# - vx * heightValue / 2#) + _
        axisY * (centerY + uy * widthValue / 2# - vy * heightValue / 2#), initialized, minValue
    TestIncludeProjectionMinimum axisX * (centerX - ux * widthValue / 2# + vx * heightValue / 2#) + _
        axisY * (centerY - uy * widthValue / 2# + vy * heightValue / 2#), initialized, minValue
    TestIncludeProjectionMinimum axisX * (centerX - ux * widthValue / 2# - vx * heightValue / 2#) + _
        axisY * (centerY - uy * widthValue / 2# - vy * heightValue / 2#), initialized, minValue
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Уточняет наименьшее значение линейной функции по очередной проверяемой точке.
Private Sub TestIncludeProjectionMinimum(ByVal candidate As Double, ByRef initialized As Boolean, _
        ByRef minValue As Double)
    If Not initialized Or candidate < minValue Then minValue = candidate
    initialized = True
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Фиксирует соглашение по справочным ядровым расстояниям главных осей.
' Для высокого симметричного прямоугольника главная ось 1 совпадает с +X,
' поэтому r по оси 1 должен совпасть с r по центральной оси X. Отдельно
' проверяем, что плоскость устойчивости 1 остается другой сущностью: ее
' нормаль направлена по оси 2, потому что изгиб идет вокруг оси 1.
Private Sub TestPrincipalAxisCoreDistanceIsAxisBased(ByRef stats As TTestStats)
    Dim section As CSectionModel
    Set section = New CSectionModel
    section.AddConcreteElement 0#, 0#, 120000#, 1, "", "", "Rectangle", 200#, 600#, 0#

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete section

    Dim axis1X As Double
    Dim axis1Y As Double
    Dim axis2X As Double
    Dim axis2Y As Double
    props.PrincipalAxisDirection 1, axis1X, axis1Y
    props.PrincipalAxisDirection 2, axis2X, axis2Y

    TestAssertClose stats, "principal.core.axis1.matchesX", _
        props.CoreDistanceAlong(section, axis1X, axis1Y, False), _
        props.CoreDistanceAlong(section, 1#, 0#, False), 0.000001
    TestAssertClose stats, "principal.core.axis2.matchesY", _
        props.CoreDistanceAlong(section, axis2X, axis2Y, False), _
        props.CoreDistanceAlong(section, 0#, 1#, False), 0.000001
    TestAssertClose stats, "principal.core.plane1.normalIsAxis2", _
        props.PrincipalPlaneCoreDistance(section, 1, True, False), _
        props.CoreDistanceAlong(section, axis2X, axis2Y, False), 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что грубая, но симметричная сетка круга не разворачивает главные
' оси из-за микроскопического численного Ixy. Для устойчивости это важно:
' чистый пользовательский Mx не должен попадать во вторую плоскость только из-за
' дискретизационного шума.
Private Sub TestCirclePrincipalAxesStableOnCoarseMesh(ByRef stats As TTestStats)
    Dim geom As CGeometryRegion
    Set geom = TestFixtureDisk(500#)

    Dim props As CSectionPropertiesCalculator
    Set props = TestMeshProps(geom, 20#, 20#)

    TestAssertClose stats, "circle.coarse.principal.angle", props.PrincipalAngleRad, 0#, 0.000000001
    TestAssertTrue stats, "circle.coarse.IxIy.nearlyEqual", _
        Abs(props.Ixc - props.Iyc) / GeomMax(Abs(props.Ixc), Abs(props.Iyc)) < 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что нулевой радиус и отрицательный диаметр не дают допустимый
' круг; геометрическая валидация выполняется до построения расчетной сетки.
Private Sub TestCircleInvalidData(ByRef stats As TTestStats)
    Dim geometry As CGeometryRegion, number As Long
    On Error Resume Next
    Set geometry = TestFixtureDisk(0#): number = Err.Number: Err.Clear
    On Error GoTo 0
    TestAssertTrue stats, "circle.invalid.zeroRadius", number <> 0 And geometry Is Nothing
    On Error Resume Next
    Set geometry = TestFixtureDisk(-100#): number = Err.Number: Err.Clear
    On Error GoTo 0
    TestAssertTrue stats, "circle.invalid.negativeDiameter", number <> 0 And geometry Is Nothing
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет сетку симметричного клиновидного сечения с аналитической
' площадью, нулевым центром/Ixy и ожидаемым направлением главных осей.
Private Sub TestSymmetricPointedSection(ByRef stats As TTestStats)
    Dim filletPolygon As CPolygonSpec
    Dim geom As CGeometryRegion
    Set filletPolygon = New CPolygonSpec
    filletPolygon.AddVertex "0", -150#, 80#, 25#
    filletPolygon.AddVertex "1", 150#, 80#, 25#
    filletPolygon.AddVertex "2", 150#, -80#, 25#
    filletPolygon.AddVertex "3", -150#, -80#, 25#
    filletPolygon.Seal
    Set geom = TestFixtureFillets(filletPolygon)

    Dim areaAvailable As Boolean
    Dim analyticalArea As Double
    analyticalArea = geom.AnalyticalArea(areaAvailable)

    Dim props As CSectionPropertiesCalculator
    Set props = TestMeshProps(geom, 2.5, 2.5)

    TestAssertTrue stats, "sym.area.available", areaAvailable
    TestAssertRelative stats, "sym.area", props.Area, analyticalArea, 0.003
    TestAssertClose stats, "sym.cx", props.CentroidX, 0#, 0.05
    TestAssertClose stats, "sym.cy", props.CentroidY, 0#, 0.05
    TestAssertClose stats, "sym.Ixy", props.Ixyc, 0#, 0.000001 * props.Area
    TestAssertClose stats, "sym.principal.angle", props.PrincipalAngleRad, GEOM_PI / 2#, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет независимую коническую сторону: сходимость сеточных A/I,
' положение касательных и фактический габарит после скруглений.
' Допустимые большие радиусы различаются с действительно перекрывающимися.
Private Sub TestTaperedPointedSection(ByRef stats As TTestStats)
    Dim filletPolygon As CPolygonSpec
    Dim polygonSpec As CPolygonSpec, polygonPoint As Variant
    Dim geom As CGeometryRegion
    Set filletPolygon = New CPolygonSpec
    filletPolygon.AddVertex "0", -150#, 90#, 20#
    filletPolygon.AddVertex "1", 150#, 90#, 30#
    filletPolygon.AddVertex "2", 230#, 0#, 45#
    filletPolygon.AddVertex "3", 150#, -90#, 30#
    filletPolygon.AddVertex "4", -150#, -90#, 20#
    filletPolygon.Seal
    Set geom = TestFixtureFillets(filletPolygon)

    Dim props25 As CSectionPropertiesCalculator, props125 As CSectionPropertiesCalculator
    Set props25 = TestMeshProps(geom, 25#, 25#): Set props125 = TestMeshProps(geom, 12.5, 12.5)
    TestAssertTrue stats, "rounded.tapered.area.positive", props125.Area > 0#
    TestAssertTrue stats, "rounded.tapered.centroid.inside", geom.ContainsPoint(props125.CentroidX, props125.CentroidY)
    TestAssertTrue stats, "rounded.tapered.rightExtension", geom.MaxX > 150# And geom.MaxX < 230#
    TestAssertClose stats, "rounded.tapered.height", geom.MaxY - geom.MinY, 180#, 0.000001
    TestAssertRelative stats, "rounded.tapered.area.convergence", props125.Area, props25.Area, 0.08
    TestAssertRelative stats, "rounded.tapered.Ix.convergence", props125.Ixc, props25.Ixc, 0.15

    Dim query As CSectionGeometryQuery, path As CGeometryPath, wide As CGeometryRegion, top As TRegionEdge
    Set query = New CSectionGeometryQuery
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "TL", -500#, 250#, 150#
    polygonSpec.AddVertex "TR", 500#, 250#, 300#
    polygonSpec.AddVertex "Nose", 750#, 0#, 100#
    polygonSpec.AddVertex "BR", 500#, -250#, 300#
    polygonSpec.AddVertex "BL", -500#, -250#, 150#
    polygonSpec.Seal
    Set path = query.RoundedPolygon("Wide", polygonSpec)
    path.CopySegment path.SegmentIndex("Edge.TL"), top
    Set wide = TestFixtureGeometry(query.RegionFromPath(path))
    TestAssertClose stats, "rounded.tapered.variableAngle.topTangent", top.X2, 500# - 300# * (Sqr(2#) - 1#), 0.000001
    TestAssertClose stats, "rounded.tapered.variableAngle.maxX", wide.MaxX, 750# - 100# * (Sqr(2#) - 1#), 0.000001
    TestAssertTrue stats, "rounded.tapered.variableAngle.maxLessThanRawApex", wide.MaxX < 750#

    Dim largeR1 As CGeometryRegion, validMessage As String
    Set filletPolygon = New CPolygonSpec
    filletPolygon.AddVertex "0", -400#, 125#, 40#
    filletPolygon.AddVertex "1", 400#, 125#, 160#
    filletPolygon.AddVertex "2", 800#, 0#, 30#
    filletPolygon.AddVertex "3", 400#, -125#, 160#
    filletPolygon.AddVertex "4", -400#, -125#, 40#
    filletPolygon.Seal
    Set largeR1 = TestFixtureFillets(filletPolygon)
    TestAssertTrue stats, "rounded.tapered.r1GreaterThanHalfHeight.valid", largeR1.IsValid(validMessage)
    Dim invalid As CGeometryRegion, number As Long, reason As String
    On Error Resume Next
    Set filletPolygon = New CPolygonSpec
    filletPolygon.AddVertex "0", -400#, 125#, 40#
    filletPolygon.AddVertex "1", 400#, 125#, 300#
    filletPolygon.AddVertex "2", 480#, 0#, 120#
    filletPolygon.AddVertex "3", 400#, -125#, 300#
    filletPolygon.AddVertex "4", -400#, -125#, 40#
    filletPolygon.Seal
    Set invalid = TestFixtureFillets(filletPolygon)
    number = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo 0
    TestAssertTrue stats, "rounded.tapered.radiiOverlap.invalid", number <> 0 And Len(reason) > 0 And invalid Is Nothing
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет semantic-аннотации PointedSection: скругления должны идти дугами,
' а размер tapered-части выводится как фактический W' до края бетона.
Private Sub TestPointedSectionContourAnnotations(ByRef stats As TTestStats)
    Dim geom As CSectionDefinition
    Dim geomParameters As Variant
    geomParameters = Array(1000#, 500#, 0#, 250#, 150#, 300#, 0#, 100#)
    Set geom = TestFixtureRoundedDefinition(geomParameters)

    Dim model As CSectionModel
    Set model = New CSectionModel
    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout

    Dim builder As CSectionAnnotationBuilder
    Set builder = New CSectionAnnotationBuilder
    geom.WriteContours model.Contours
    builder.Build geom, rebars, model.Annotations

    Dim annotations As CSectionAnnotations
    Set annotations = model.Annotations

    TestAssertTrue stats, "rounded.annotation.hasArcs", TestCountSectionAnnotationType(model, "CONTOUR_ARC") > 0
    TestAssertTrue stats, "rounded.annotation.hasLines", TestCountSectionAnnotationType(model, "CONTOUR_LINE") > 0

    Dim wIndex As Long
    wIndex = TestFindSectionAnnotationIndex(model, "DIMENSION", "DIM_W_RIGHT")
    TestAssertTrue stats, "rounded.annotation.wRight.exists", wIndex > 0
    If wIndex > 0 Then
        TestAssertTrue stats, "rounded.annotation.wRight.label", annotations.Text(wIndex) = "W'"
        TestAssertClose stats, "rounded.annotation.wRight.value", annotations.Value(wIndex), _
            TestRoundedDisplayLength(TestFixtureGeometry(geom.ConcreteDomain).MaxX - 500#), 0.000001
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяет пользовательское округление справочного размера W' до целых мм.
Private Function TestRoundedDisplayLength(ByVal valueMm As Double) As Double
    If valueMm >= 0# Then
        TestRoundedDisplayLength = Int(valueMm + 0.5)
    Else
        TestRoundedDisplayLength = -Int(Abs(valueMm) + 0.5)
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет новую автоматическую раскладку PointedSection: горизонтальные
' линии B идут только по прямому участку между R1, а боковая H-линия может
' идти по tapered-траектории.
Private Sub TestPointedSectionRebarLayout(ByRef stats As TTestStats)
    Dim geom As CSectionDefinition
    Dim geomParameters As Variant
    geomParameters = Array(800#, 250#, 0#, 150#, 40#, 40#, 0#, 80#)
    Set geom = TestFixtureRoundedDefinition(geomParameters)

    Dim builder As CRebarRuleExecutor
    Set builder = New CRebarRuleExecutor

    Dim layout As CRebarLayout
    Set geom = TestFixtureRoundedDefinition(geomParameters, _
        Array(Array(40#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 5, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 5, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")))
    Set layout = TestFixtureAnnotatedLayout(geom, builder)

    TestAssertTrue stats, "rounded.rebar.count", layout.Count = 16
    TestAssertTrue stats, "rounded.rebar.requested", builder.RequestedPositions = 16
    TestAssertClose stats, "rounded.rebar.top.firstX", layout.X(1), -360#, 0.000001
    TestAssertClose stats, "rounded.rebar.top.firstY", layout.Y(1), 85#, 0.000001
    Dim topX1 As Double
    Dim topY1 As Double
    Dim topX2 As Double
    Dim topY2 As Double
    Dim normalX As Double
    Dim normalY As Double
    TestFixturePathEnds geom.Path("B.Top"), topX1, topY1, topX2, topY2, normalX, normalY
    TestAssertClose stats, "rounded.rebar.top.lastX", layout.X(5), topX2, 0.000001
    TestAssertTrue stats, "rounded.rebar.top.lastX.followsTaperAngle", layout.X(5) > 360# And layout.X(5) < 400#
    TestAssertClose stats, "rounded.rebar.bottom.firstY", layout.Y(6), -85#, 0.000001

    Dim bottomX1 As Double
    Dim bottomY1 As Double
    Dim bottomX2 As Double
    Dim bottomY2 As Double
    TestFixturePathEnds geom.Path("B.Bottom"), bottomX1, bottomY1, bottomX2, bottomY2, normalX, normalY

    Dim leftPath As CGeometryPath, rightPath As CGeometryPath
    Dim leftStartX As Double, leftStartY As Double, leftEndX As Double, leftEndY As Double
    Set leftPath = geom.Path("H.Left")
    TestFixturePathEnds leftPath, leftStartX, leftStartY, leftEndX, leftEndY, normalX, normalY
    TestAssertClose stats, "rounded.rebar.hLeft.pathStartX", leftStartX, topX1, 0.000001
    TestAssertClose stats, "rounded.rebar.hLeft.pathStartY", leftStartY, topY1, 0.000001
    TestAssertClose stats, "rounded.rebar.hLeft.pathEndX", leftEndX, bottomX1, 0.000001
    TestAssertClose stats, "rounded.rebar.hLeft.pathEndY", leftEndY, bottomY1, 0.000001

    Dim leftStep As Double
    leftStep = leftPath.Length / 4#
    Dim expectedX As Double
    Dim expectedY As Double
    leftPath.AtDistance leftStep, expectedX, expectedY, normalX, normalY
    TestAssertClose stats, "rounded.rebar.hLeft.firstAtOneStepX", layout.X(11), expectedX, 0.000001
    TestAssertClose stats, "rounded.rebar.hLeft.firstAtOneStepY", layout.Y(11), expectedY, 0.000001
    leftPath.AtDistance leftStep * 3#, expectedX, expectedY, normalX, normalY
    TestAssertClose stats, "rounded.rebar.hLeft.lastAtOneStepX", layout.X(13), expectedX, 0.000001
    TestAssertClose stats, "rounded.rebar.hLeft.lastAtOneStepY", layout.Y(13), expectedY, 0.000001

    Dim rightStartX As Double, rightStartY As Double, rightEndX As Double, rightEndY As Double
    Set rightPath = geom.Path("H.Right")
    TestFixturePathEnds rightPath, rightStartX, rightStartY, rightEndX, rightEndY, normalX, normalY
    TestAssertClose stats, "rounded.rebar.hRight.pathStartX", rightStartX, topX2, 0.000001
    TestAssertClose stats, "rounded.rebar.hRight.pathStartY", rightStartY, topY2, 0.000001
    TestAssertClose stats, "rounded.rebar.hRight.pathEndX", rightEndX, bottomX2, 0.000001
    TestAssertClose stats, "rounded.rebar.hRight.pathEndY", rightEndY, bottomY2, 0.000001
    Dim rightStep As Double
    rightStep = rightPath.Length / 4#
    rightPath.AtDistance rightStep, expectedX, expectedY, normalX, normalY
    TestAssertClose stats, "rounded.rebar.hRight.firstAtOneStepX", layout.X(14), expectedX, 0.000001
    TestAssertClose stats, "rounded.rebar.hRight.firstAtOneStepY", layout.Y(14), expectedY, 0.000001

    Dim simpleGeom As CSectionDefinition
    Dim simpleGeomParameters As Variant
    simpleGeomParameters = Array(800#, 250#, 0#, 0#, 40#, 40#, 0#, 0#)
    Dim simpleFace As Variant
    simpleFace = Array(20#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
    Set simpleGeom = TestFixtureRoundedDefinition(simpleGeomParameters, Array(simpleFace, simpleFace, simpleFace, simpleFace))

    Dim simpleTopX1 As Double
    Dim simpleTopY1 As Double
    Dim simpleTopX2 As Double
    Dim simpleTopY2 As Double
    Dim simpleBottomX1 As Double
    Dim simpleBottomY1 As Double
    Dim simpleBottomX2 As Double
    Dim simpleBottomY2 As Double
    TestFixturePathEnds simpleGeom.Path("B.Top"), simpleTopX1, simpleTopY1, simpleTopX2, simpleTopY2, normalX, normalY
    TestFixturePathEnds simpleGeom.Path("B.Bottom"), simpleBottomX1, simpleBottomY1, simpleBottomX2, simpleBottomY2, normalX, normalY

    Dim simpleLeftStartX As Double, simpleLeftStartY As Double, simpleLeftEndX As Double, simpleLeftEndY As Double
    TestFixturePathEnds simpleGeom.Path("H.Left"), simpleLeftStartX, simpleLeftStartY, simpleLeftEndX, simpleLeftEndY, normalX, normalY
    TestAssertClose stats, "rounded.rebar.simpleLeft.pathStartX", simpleLeftStartX, simpleTopX1, 0.000001
    TestAssertClose stats, "rounded.rebar.simpleLeft.pathStartY", simpleLeftStartY, simpleTopY1, 0.000001
    TestAssertClose stats, "rounded.rebar.simpleLeft.pathEndX", simpleLeftEndX, simpleBottomX1, 0.000001
    TestAssertClose stats, "rounded.rebar.simpleLeft.pathEndY", simpleLeftEndY, simpleBottomY1, 0.000001

    Dim simpleRightStartX As Double, simpleRightStartY As Double, simpleRightEndX As Double, simpleRightEndY As Double
    TestFixturePathEnds simpleGeom.Path("H.Right"), simpleRightStartX, simpleRightStartY, simpleRightEndX, simpleRightEndY, normalX, normalY
    TestAssertClose stats, "rounded.rebar.simpleRight.pathStartX", simpleRightStartX, simpleTopX2, 0.000001
    TestAssertClose stats, "rounded.rebar.simpleRight.pathStartY", simpleRightStartY, simpleTopY2, 0.000001
    TestAssertClose stats, "rounded.rebar.simpleRight.pathEndX", simpleRightEndX, simpleBottomX2, 0.000001
    TestAssertClose stats, "rounded.rebar.simpleRight.pathEndY", simpleRightEndY, simpleBottomY2, 0.000001

    Dim i As Long
    For i = 1 To layout.Count
        TestAssertTrue stats, "rounded.rebar.inside." & CStr(i), TestFixtureGeometry(geom.ConcreteDomain).ContainsPoint(layout.X(i), layout.Y(i))
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет базовую математику HollowRectangle: бетонная область равна
' наружному PointedSection минус Opening, а точки внутри отверстия не входят
' в бетонную часть сечения.
Private Sub TestHollowRectangleGeometry(ByRef stats As TTestStats)
    Dim polygonSpec As CPolygonSpec, polygonPoint As Variant
    Dim geom As CGeometryRegion, facade As CSectionDefinitionBuilder, path As CGeometryPath, definition As CSectionDefinition
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "HollowGeometryFixture"
    Set path = facade.RectanglePath("Boundary", -250#, -400#, 500#, 800#): facade.Outer "OUTER", path.ID
    Set path = facade.RectanglePath("Void", -100#, -250#, 200#, 500#): facade.Opening "VOID", path.ID, "OUTER"
    Set definition = facade.Finish: Set geom = TestFixtureGeometry(definition.ConcreteDomain)

    Dim message As String
    TestAssertTrue stats, "hollow.geometry.valid", geom.IsValid(message)

    Dim available As Boolean
    Dim areaValue As Double
    areaValue = geom.AnalyticalArea(available)
    TestAssertTrue stats, "hollow.geometry.areaAvailable", available
    TestAssertClose stats, "hollow.geometry.area", areaValue, 500# * 800# - 200# * 500#, 0.000001

    Dim cx As Double
    Dim cy As Double
    geom.AnalyticalCentroid available, cx, cy
    TestAssertClose stats, "hollow.geometry.cx", cx, 0#, 0.000001
    TestAssertClose stats, "hollow.geometry.cy", cy, 0#, 0.000001

    TestAssertTrue stats, "hollow.geometry.outerConcrete", geom.ContainsPoint(0#, 350#)
    TestAssertTrue stats, "hollow.geometry.openingExcluded", Not geom.ContainsPoint(0#, 0#)
    TestAssertTrue stats, "hollow.geometry.outsideExcluded", Not geom.ContainsPoint(260#, 0#)

    Dim offsetGeom As CGeometryRegion
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "OffsetVoidFixture"
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", -250#, -400#, 30#
    polygonSpec.AddVertex "BR", 250#, -400#, 30#
    polygonSpec.AddVertex "TR", 250#, 400#, 30#
    polygonSpec.AddVertex "TL", -250#, 400#, 30#
    polygonSpec.Seal
    Set path = facade.RoundedPolygon("Boundary", polygonSpec)
    facade.Outer "OUTER", path.ID
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", -60#, -180#, 20#
    polygonSpec.AddVertex "BR", 140#, -180#, 20#
    polygonSpec.AddVertex "TR", 140#, 120#, 20#
    polygonSpec.AddVertex "TL", -60#, 120#, 20#
    polygonSpec.Seal
    Set path = facade.RoundedPolygon("Void", polygonSpec)
    facade.Opening "VOID", path.ID, "OUTER"
    Set definition = facade.Finish: Set offsetGeom = TestFixtureGeometry(definition.ConcreteDomain)
    areaValue = offsetGeom.AnalyticalArea(available)
    offsetGeom.AnalyticalCentroid available, cx, cy
    TestAssertTrue stats, "hollow.geometry.offset.valid", offsetGeom.IsValid(message)
    TestAssertTrue stats, "hollow.geometry.offset.cxOppositeOpening", cx < 0#
    TestAssertTrue stats, "hollow.geometry.offset.cyOppositeOpening", cy > 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет semantic-контуры HollowRectangle. Наружная граница и Opening должны
' сохраняться разными петлями, чтобы схема и AutoCAD export не соединяли их
' случайной линией.
Private Sub TestHollowRectangleContourAnnotations(ByRef stats As TTestStats)
    Dim geom As CSectionDefinition
    Dim geomParameters As Variant
    geomParameters = Array(500#, 800#, 180#, 200#, 500#, 30#)
    Set geom = TestFixtureHollowDefinition(geomParameters)

    Dim model As CSectionModel
    Set model = New CSectionModel
    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout

    Dim builder As CSectionAnnotationBuilder
    Set builder = New CSectionAnnotationBuilder
    geom.WriteContours model.Contours
    builder.Build geom, rebars, model.Annotations

    TestAssertTrue stats, "hollow.annotation.outerLoop", TestHasSectionAnnotationIDPrefix(model, "CONTOUR_OUTER_")
    TestAssertTrue stats, "hollow.annotation.openingLoop", TestHasSectionAnnotationIDPrefix(model, "CONTOUR_OPENING_")
    TestAssertTrue stats, "hollow.annotation.arcs", TestCountSectionAnnotationType(model, "CONTOUR_ARC") >= 8
    TestAssertTrue stats, "hollow.annotation.dimB", TestHasSectionAnnotation(model, "DIMENSION", "DIM_B")
    TestAssertTrue stats, "hollow.annotation.dimOpeningB", TestHasSectionAnnotation(model, "DIMENSION", "DIM_B_OPENING")
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет раскладку HollowRectangle: внешние грани получают пользовательское
' количество стержней, а внутренние грани Opening заполняются проекциями
' соответствующих внешних граней.
Private Sub TestHollowRectangleRebarLayout(ByRef stats As TTestStats)
    Dim geom As CSectionDefinition
    Dim geomParameters As Variant
    geomParameters = Array(500#, 800#, 180#, 200#, 500#, 30#)
    Set geom = TestFixtureHollowDefinition(geomParameters)

    Dim builder As CRebarRuleExecutor
    Set builder = New CRebarRuleExecutor

    Dim layout As CRebarLayout, annotations As CSectionAnnotations
    Set geom = TestFixtureHollowDefinition(geomParameters, _
        Array(Array(40#, 20#, 6, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 6, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 8, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 8, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 0, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 0, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 0, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 0, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")))
    Set layout = TestFixtureAnnotatedLayout(geom, builder, annotations)

    TestAssertTrue stats, "hollow.rebar.external.top", TestCountBarsInAnnotationGroup(layout, "B.Top") = 8
    TestAssertTrue stats, "hollow.rebar.external.bottom", TestCountBarsInAnnotationGroup(layout, "B.Bottom") = 8
    TestAssertTrue stats, "hollow.rebar.external.left", TestCountBarsInAnnotationGroup(layout, "H.Left") = 6
    TestAssertTrue stats, "hollow.rebar.external.right", TestCountBarsInAnnotationGroup(layout, "H.Right") = 6

    Dim hLeftX1 As Double, hLeftY1 As Double, hLeftX2 As Double, hLeftY2 As Double
    Dim hLeftNormalX As Double, hLeftNormalY As Double
    TestFixturePathEnds geom.Path("H.Left"), hLeftX1, hLeftY1, hLeftX2, hLeftY2, hLeftNormalX, hLeftNormalY
    Dim firstHLeft As Long, lastHLeft As Long
    TestFindFirstLastBarInAnnotationGroup layout, "H.Left", firstHLeft, lastHLeft
    TestAssertClose stats, "hollow.rebar.hLeft.firstX", layout.X(firstHLeft), hLeftX1, 0.000001
    TestAssertClose stats, "hollow.rebar.hLeft.firstY", layout.Y(firstHLeft), hLeftY1, 0.000001
    TestAssertClose stats, "hollow.rebar.hLeft.lastX", layout.X(lastHLeft), hLeftX2, 0.000001
    TestAssertClose stats, "hollow.rebar.hLeft.lastY", layout.Y(lastHLeft), hLeftY2, 0.000001

    Dim annotationBTop As Long
    annotationBTop = TestFindRebarAnnotationAnchor(annotations, "B.Top")
    Dim bTopMaxY As Double
    bTopMaxY = TestExtremeYInAnnotationGroup(layout, "B.Top", True)
    TestAssertClose stats, "hollow.rebar.bTop.anchorStartY", annotations.StartY(annotationBTop), bTopMaxY, 0.000001
    TestAssertClose stats, "hollow.rebar.bTop.anchorEndY", annotations.EndY(annotationBTop), bTopMaxY, 0.000001

    Dim annotationBBottom As Long
    annotationBBottom = TestFindRebarAnnotationAnchor(annotations, "B.Bottom")
    Dim bBottomMinY As Double
    bBottomMinY = TestExtremeYInAnnotationGroup(layout, "B.Bottom", False)
    TestAssertClose stats, "hollow.rebar.bBottom.anchorStartY", annotations.StartY(annotationBBottom), bBottomMinY, 0.000001
    TestAssertClose stats, "hollow.rebar.bBottom.anchorEndY", annotations.EndY(annotationBBottom), bBottomMinY, 0.000001

    Dim annotationOpeningBTop As Long
    annotationOpeningBTop = TestFindRebarAnnotationAnchor(annotations, "Opening.B.Top")
    Dim openingBTopMaxY As Double
    openingBTopMaxY = TestExtremeYInAnnotationGroup(layout, "Opening.B.Top", True)
    TestAssertClose stats, "hollow.rebar.openingBTop.anchorStartY", annotations.StartY(annotationOpeningBTop), openingBTopMaxY, 0.000001
    TestAssertClose stats, "hollow.rebar.openingBTop.anchorEndY", annotations.EndY(annotationOpeningBTop), openingBTopMaxY, 0.000001

    Dim annotationOpeningBBottom As Long
    annotationOpeningBBottom = TestFindRebarAnnotationAnchor(annotations, "Opening.B.Bottom")
    Dim openingBBottomMinY As Double
    openingBBottomMinY = TestExtremeYInAnnotationGroup(layout, "Opening.B.Bottom", False)
    TestAssertClose stats, "hollow.rebar.openingBBottom.anchorStartY", annotations.StartY(annotationOpeningBBottom), openingBBottomMinY, 0.000001
    TestAssertClose stats, "hollow.rebar.openingBBottom.anchorEndY", annotations.EndY(annotationOpeningBBottom), openingBBottomMinY, 0.000001

    TestAssertTrue stats, "hollow.rebar.opening.topProjected", TestCountBarsInAnnotationGroup(layout, "Opening.B.Top") > 0
    TestAssertTrue stats, "hollow.rebar.opening.bottomProjected", TestCountBarsInAnnotationGroup(layout, "Opening.B.Bottom") > 0
    TestAssertTrue stats, "hollow.rebar.opening.leftProjected", TestCountBarsInAnnotationGroup(layout, "Opening.H.Left") > 0
    TestAssertTrue stats, "hollow.rebar.opening.rightProjected", TestCountBarsInAnnotationGroup(layout, "Opening.H.Right") > 0

    Dim i As Long
    For i = 1 To layout.Count
        TestAssertTrue stats, "hollow.rebar.inside." & CStr(i), TestFixtureGeometry(geom.ConcreteDomain).ContainsPoint(layout.X(i), layout.Y(i))
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Фиксирует прямоугольный случай R=R_o=0: внутренняя B-линия Opening должна
' быть прямой offset-линией, продленной до боковых offset-линий H.Left/H.Right.
' Крайние проекции не должны падать на вертикальные стороны Opening и уходить
' по Y от заданного as_B.
Private Sub TestHollowRectangleSharpOpeningBProjection(ByRef stats As TTestStats)
    Dim geom As CSectionDefinition
    Dim geomParameters As Variant
    geomParameters = Array(500#, 800#, 0#, 200#, 500#, 0#)
    Set geom = TestFixtureHollowDefinition(geomParameters)

    Dim builder As CRebarRuleExecutor
    Set builder = New CRebarRuleExecutor

    Dim layout As CRebarLayout
    Set geom = TestFixtureHollowDefinition(geomParameters, _
        Array(Array(40#, 20#, 6, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 6, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 8, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 8, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 0, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 0, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 0, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), Array(40#, 20#, 0, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")))
    Set layout = TestFixtureAnnotatedLayout(geom, builder)

    TestAssertTrue stats, "hollow.rebar.sharp.openingTop.count", TestCountBarsInAnnotationGroup(layout, "Opening.B.Top") = 6
    TestAssertTrue stats, "hollow.rebar.sharp.openingBottom.count", TestCountBarsInAnnotationGroup(layout, "Opening.B.Bottom") = 6
    TestAssertTrue stats, "hollow.rebar.sharp.openingLeft.count", TestCountBarsInAnnotationGroup(layout, "Opening.H.Left") = 4
    TestAssertTrue stats, "hollow.rebar.sharp.openingRight.count", TestCountBarsInAnnotationGroup(layout, "Opening.H.Right") = 4

    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim normalX As Double, normalY As Double
    TestFixturePathEnds geom.Path("Opening.B.Top"), x1, y1, x2, y2, normalX, normalY
    TestAssertBarsInGroupHaveY stats, layout, "Opening.B.Top", y1, "hollow.rebar.sharp.openingTop"

    TestFixturePathEnds geom.Path("Opening.B.Bottom"), x1, y1, x2, y2, normalX, normalY
    TestAssertBarsInGroupHaveY stats, layout, "Opening.B.Bottom", y1, "hollow.rebar.sharp.openingBottom"

    TestFixturePathEnds geom.Path("Opening.H.Left"), x1, y1, x2, y2, normalX, normalY
    TestAssertBarsInGroupHaveX stats, layout, "Opening.H.Left", x1, "hollow.rebar.sharp.openingLeft"

    TestFixturePathEnds geom.Path("Opening.H.Right"), x1, y1, x2, y2, normalX, normalY
    TestAssertBarsInGroupHaveX stats, layout, "Opening.H.Right", x1, "hollow.rebar.sharp.openingRight"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет ошибки размеров/радиусов клиновидного сечения и неположительных
' шагов сетки. Невалидная геометрия не должна создавать расчетные волокна.
Private Sub TestInvalidData(ByRef stats As TTestStats)
    TestAssertInvalid stats, "invalid.width", -100#, 100#, 0#, 0#, 0#, 0#
    TestAssertInvalid stats, "invalid.height", 100#, 0#, 0#, 0#, 0#, 0#
    TestAssertInvalid stats, "invalid.negative.radius", 100#, 100#, -1#, 0#, 0#, 0#
    TestAssertInvalid stats, "invalid.large.radius", 100#, 100#, 101#, 0#, 0#, 0#
    TestAssertInvalid stats, "invalid.neighbor.radii", 100#, 100#, 60#, 60#, 0#, 0#

    Dim geom As CGeometryRegion
    Set geom = TestFixtureBox(100#, 100#)
    TestAssertBuildError stats, "invalid.step.zero", geom, 0#, 10#
    TestAssertBuildError stats, "invalid.step.negative", geom, 10#, -10#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит одно скругленное сечение с шагами 50/25/12.5 мм и записывает A/центр/I.
' Последнее уточнение проверяется прежними относительными допусками;
' численные результаты не заменяются текущей сеткой как новым эталоном.
Private Sub TestMeshConvergence(ByRef stats As TTestStats)
    Dim filletPolygon As CPolygonSpec
    Dim geom As CGeometryRegion
    Set filletPolygon = New CPolygonSpec
    filletPolygon.AddVertex "0", -200#, 110#, 35#
    filletPolygon.AddVertex "1", 200#, 110#, 35#
    filletPolygon.AddVertex "2", 200#, -110#, 35#
    filletPolygon.AddVertex "3", -200#, -110#, 35#
    filletPolygon.Seal
    Set geom = TestFixtureFillets(filletPolygon)

    Dim p50 As CSectionPropertiesCalculator
    Dim p25 As CSectionPropertiesCalculator
    Dim p125 As CSectionPropertiesCalculator

    Set p50 = TestMeshProps(geom, 50#, 50#)
    Set p25 = TestMeshProps(geom, 25#, 25#)
    Set p125 = TestMeshProps(geom, 12.5, 12.5)

    TestAppendLine stats, "CONVERGENCE: step=50 area=" & TestFormatNumberInvariant(p50.Area) & _
        "; cx=" & TestFormatNumberInvariant(p50.CentroidX) & "; cy=" & TestFormatNumberInvariant(p50.CentroidY) & _
        "; Ix=" & TestFormatNumberInvariant(p50.Ixc) & "; Iy=" & TestFormatNumberInvariant(p50.Iyc) & _
        "; Ixy=" & TestFormatNumberInvariant(p50.Ixyc) & "; I1=" & TestFormatNumberInvariant(p50.PrincipalI1) & _
        "; I2=" & TestFormatNumberInvariant(p50.PrincipalI2)
    TestAppendLine stats, "CONVERGENCE: step=25 area=" & TestFormatNumberInvariant(p25.Area) & _
        "; cx=" & TestFormatNumberInvariant(p25.CentroidX) & "; cy=" & TestFormatNumberInvariant(p25.CentroidY) & _
        "; Ix=" & TestFormatNumberInvariant(p25.Ixc) & "; Iy=" & TestFormatNumberInvariant(p25.Iyc) & _
        "; Ixy=" & TestFormatNumberInvariant(p25.Ixyc) & "; I1=" & TestFormatNumberInvariant(p25.PrincipalI1) & _
        "; I2=" & TestFormatNumberInvariant(p25.PrincipalI2)
    TestAppendLine stats, "CONVERGENCE: step=12.5 area=" & TestFormatNumberInvariant(p125.Area) & _
        "; cx=" & TestFormatNumberInvariant(p125.CentroidX) & "; cy=" & TestFormatNumberInvariant(p125.CentroidY) & _
        "; Ix=" & TestFormatNumberInvariant(p125.Ixc) & "; Iy=" & TestFormatNumberInvariant(p125.Iyc) & _
        "; Ixy=" & TestFormatNumberInvariant(p125.Ixyc) & "; I1=" & TestFormatNumberInvariant(p125.PrincipalI1) & _
        "; I2=" & TestFormatNumberInvariant(p125.PrincipalI2)

    TestAssertRelative stats, "conv.area.25.12", p125.Area, p25.Area, 0.03
    TestAssertRelative stats, "conv.Ix.25.12", p125.Ixc, p25.Ixc, 0.05
    TestAssertRelative stats, "conv.Iy.25.12", p125.Iyc, p25.Iyc, 0.05
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Измеряет построение сетки и геометрические характеристики модели 800x500.
' Проверяет непустую сетку; время записывается как диагностика, без
' необоснованного требования ускорения отдельного случайного запуска.
Private Sub TestPerformance(ByRef stats As TTestStats)
    Dim filletPolygon As CPolygonSpec
    Dim geom As CGeometryRegion
    Set filletPolygon = New CPolygonSpec
    filletPolygon.AddVertex "0", -400#, 250#, 40#
    filletPolygon.AddVertex "1", 400#, 250#, 40#
    filletPolygon.AddVertex "2", 400#, -250#, 40#
    filletPolygon.AddVertex "3", -400#, -250#, 40#
    filletPolygon.Seal
    Set geom = TestFixtureFillets(filletPolygon)

    Dim builder As CFiberMeshBuilder
    Set builder = New CFiberMeshBuilder
    builder.BuildMesh geom, 10#, 10#, 1

    Dim calc As CSectionPropertiesCalculator
    Set calc = New CSectionPropertiesCalculator
    calc.CalculateConcrete TestBuildGeneratedSectionModel(builder, Nothing)

    TestAppendLine stats, "PERFORMANCE: section=800x500; step=10; fibers=" & CStr(builder.FiberCount) & _
        "; buildSec=" & TestFormatNumberInvariant(builder.TestBuildSeconds) & _
        "; propsSec=" & TestFormatNumberInvariant(calc.TestLastCalculationSeconds) & _
        "; area=" & TestFormatNumberInvariant(calc.Area)
    TestAssertTrue stats, "perf.fibers.positive", builder.FiberCount > 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Формирует локальные Ix/Iy/Ixy после поворота главных осей элемента.
' minorInertia относится к оси вдоль локальной стороны width, majorInertia -
' к перпендикулярной главной оси.
Private Sub TestRotatedLocalInertia(ByVal minorInertia As Double, ByVal majorInertia As Double, _
        ByVal angle As Double, ByRef localIx As Double, ByRef localIy As Double, _
        ByRef localIxy As Double)
    Dim c As Double
    Dim s As Double
    c = Cos(angle)
    s = Sin(angle)
    localIx = minorInertia * c * c + majorInertia * s * s
    localIy = minorInertia * s * s + majorInertia * c * c
    localIxy = (majorInertia - minorInertia) * s * c
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создает минимальный fake AutoCAD Region для тестов live-пути importer-а.
Private Function TestFakeRegion(ByVal area As Double, ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal localIx As Double, ByVal localIy As Double, ByVal localIxy As Double, _
        ByVal layerName As String, ByVal handleText As String, _
        Optional ByVal edgeRotation As Double = 0#, Optional ByVal hasEdge As Boolean = False) As CTestFakeAcadRegion
    Dim region As CTestFakeAcadRegion
    Set region = New CTestFakeAcadRegion
    region.TestInitialize area, xCoord, yCoord, localIx, localIy, localIxy, _
        layerName, handleText, edgeRotation, hasEdge
    Set TestFakeRegion = region
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит сетку общими геометрическими классами и вычисляет площадь, центр
' и моменты инерции бетона. В ошибке сохраняет шаги сетки и число элементов.
Private Function TestMeshProps(ByVal geom As CGeometryRegion, ByVal stepX As Double, ByVal stepY As Double) As CSectionPropertiesCalculator
    On Error GoTo Failed

    Dim builder As CFiberMeshBuilder
    Set builder = New CFiberMeshBuilder
    builder.BuildMesh geom, stepX, stepY, 1
    If builder.FiberCount <= 0 Then
        Err.Raise vbObjectError + 2310, "MeshProps", "Mesh is empty; stepX=" & CStr(stepX) & "; stepY=" & CStr(stepY)
    End If

    Dim calc As CSectionPropertiesCalculator
    Set calc = New CSectionPropertiesCalculator
    calc.CalculateConcrete TestBuildGeneratedSectionModel(builder, Nothing)
    Set TestMeshProps = calc
    Exit Function

Failed:
    Dim fiberCountText As String
    If builder Is Nothing Then
        fiberCountText = "builder is Nothing"
    Else
        fiberCountText = "fibers=" & CStr(builder.FiberCount)
    End If
    Err.Raise Err.Number, "MeshProps", Err.Description & "; stepX=" & CStr(stepX) & "; stepY=" & CStr(stepY) & "; " & fiberCountText
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выявляет подячейки границы хотя бы по одному размеру относительно базового шага.
Private Function TestMeshHasSmallFibers(ByVal mesh As CFiberMeshBuilder, ByVal baseStep As Double) As Boolean
    Dim i As Long
    For i = 1 To mesh.FiberCount
        If mesh.FiberWidth(i) < baseStep Or mesh.FiberHeight(i) < baseStep Then
            TestMeshHasSmallFibers = True
            Exit Function
        End If
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает ширину первой уменьшенной по X подячейки; ноль означает отсутствие такой ячейки.
Private Function TestFirstSmallFiberWidth(ByVal mesh As CFiberMeshBuilder, ByVal baseStep As Double) As Double
    Dim i As Long
    For i = 1 To mesh.FiberCount
        If mesh.FiberWidth(i) < baseStep Then
            TestFirstSmallFiberWidth = mesh.FiberWidth(i)
            Exit Function
        End If
    Next i
    TestFirstSmallFiberWidth = 0#
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отказ при недопустимых размерах и радиусах. После первой ошибки
' дальнейшее построение прекращается, чтобы не проверить другую фигуру без отвергнутой вершины.
Private Sub TestAssertInvalid(ByRef stats As TTestStats, ByVal name As String, _
        ByVal width As Double, ByVal height As Double, ByVal rtl As Double, ByVal rtr As Double, _
        ByVal rbr As Double, ByVal rbl As Double)
    Dim filletPolygon As CPolygonSpec
    Dim geometry As CGeometryRegion, number As Long
    On Error GoTo Rejected
    If width <= 0# Or height <= 0# Then
        Set geometry = TestFixtureGeometry(TestFixtureRoundedDefinition(Array(width, height, _
            0#, 0#, rtl, rtr, 0#, 0#)).ConcreteDomain)
    Else
        Set filletPolygon = New CPolygonSpec
        filletPolygon.AddVertex "0", -width / 2#, height / 2#, rtl
        filletPolygon.AddVertex "1", width / 2#, height / 2#, rtr
        filletPolygon.AddVertex "2", width / 2#, -height / 2#, rbr
        filletPolygon.AddVertex "3", -width / 2#, -height / 2#, rbl
        filletPolygon.Seal
        Set geometry = TestFixtureFillets(filletPolygon)
    End If
    On Error GoTo 0
    TestAssertTrue stats, name, False
    Exit Sub
Rejected:
    number = Err.Number
    Err.Clear
    On Error GoTo 0
    TestAssertTrue stats, name, number <> 0 And geometry Is Nothing
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Требует отклонения ошибочной геометрии/шага публичным BuildMesh, а не пустой успешной сетки.
Private Sub TestAssertBuildError(ByRef stats As TTestStats, ByVal name As String, _
        ByVal geom As CGeometryRegion, ByVal stepX As Double, ByVal stepY As Double)

    On Error GoTo GotError
    Dim builder As CFiberMeshBuilder
    Set builder = New CFiberMeshBuilder
    builder.BuildMesh geom, stepX, stepY, 1
    On Error GoTo 0
    TestAssertTrue stats, name, False
    Exit Sub

GotError:
    On Error GoTo 0
    TestAssertTrue stats, name, True
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Пропускает ошибочные круговые ряды через реальный builder и требует ошибки ввода.
Private Sub TestAssertCircleRebarError(ByRef stats As TTestStats, ByVal name As String, _
        ByVal diameter As Double, ByVal axisDistance As Double, ByVal barCount As Long, ByVal barDiameter As Double, _
        Optional ByVal row2Diameter As Double = 0#, Optional ByVal row3Diameter As Double = 0#, _
        Optional ByVal row2Location As String = "Stacked", Optional ByVal row3Location As String = "Stacked")

    On Error GoTo GotError
    Dim builder As CRebarRuleExecutor
    Set builder = New CRebarRuleExecutor
    Dim layout As CRebarLayout
    Set layout = TestFixtureMoveLayout(TestFixtureAnnotatedLayout(TestFixtureCircleDefinition(Array(diameter, axisDistance, barCount, barDiameter, row2Diameter, row3Diameter, row2Location, row3Location)), builder), 0#, 0#)
    On Error GoTo 0
    TestAssertTrue stats, name, False
    Exit Sub

GotError:
    On Error GoTo 0
    TestAssertTrue stats, name, True
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Активирует одну грань Rectangles с заданными n/as/d/отступами; остальные грани выключены.
' Ошибочные активные параметры должны отклоняться builder-ом.
Private Sub TestAssertRectanglesRebarError(ByRef stats As TTestStats, ByVal name As String, _
        ByVal lowerWidth As Double, ByVal upperHeight As Double, ByVal upperWidth As Double, ByVal lowerHeight As Double, _
        ByVal barCount As Long, ByVal axisDistance As Double, ByVal barDiameter As Double, _
        Optional ByVal startOffset As Double = 100#, Optional ByVal endOffset As Double = 100#)

    On Error GoTo GotError
    Dim builder As CRebarRuleExecutor
    Set builder = New CRebarRuleExecutor
    Dim layout As CRebarLayout
    Set layout = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(upperWidth, upperHeight, lowerWidth, lowerHeight, 0#, "TwoRect"), _
        Array(Array(axisDistance, axisDistance, barDiameter, barDiameter, barCount, 0, startOffset, endOffset, startOffset, endOffset), Array(axisDistance, axisDistance, barDiameter, barDiameter, 0, 0, startOffset, endOffset, startOffset, endOffset), Array(axisDistance, axisDistance, barDiameter, barDiameter, 0, 0, startOffset, endOffset, startOffset, endOffset), Array(axisDistance, axisDistance, barDiameter, barDiameter, 0, 0, startOffset, endOffset, startOffset, endOffset))), builder)
    On Error GoTo 0
    TestAssertTrue stats, name, False
    Exit Sub

GotError:
    On Error GoTo 0
    TestAssertTrue stats, name, True
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Фиксированные параметры выключенной грани: оба количества равны нулю, ряды отсутствуют.
Private Function TestEmptyFaceSettings() As Variant
    TestEmptyFaceSettings = Array(40#, 40#, 20#, 20#, 0, 0, 40#, 40#, 40#, 40#, 0#, 0#, 0#, 0#, "Stacked", "Stacked")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет наличие semantic-группы раскладки независимо от графического вывода.
Private Function TestHasAnnotationGroup(ByVal annotations As CSectionAnnotations, ByVal groupName As String) As Boolean
    TestHasAnnotationGroup = (TestFindRebarAnnotationAnchor(annotations, groupName) > 0)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Поиск аннотации требует совпадения типа и ID, а не только похожей подписи.
Private Function TestHasSectionAnnotation(ByVal model As CSectionModel, ByVal annotationType As String, _
        ByVal annotationID As String) As Boolean
    TestHasSectionAnnotation = (TestFindSectionAnnotationIndex(model, annotationType, annotationID) > 0)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет происхождение semantic-аннотации по префиксу ID в готовой модели.
Private Function TestHasSectionAnnotationIDPrefix(ByVal model As CSectionModel, ByVal annotationIDPrefix As String) As Boolean
    If model Is Nothing Then Exit Function

    Dim annotations As CSectionContours
    Set annotations = model.Contours

    Dim i As Long
    For i = 1 To annotations.Count
        If InStr(1, annotations.ContourID(i), annotationIDPrefix, vbTextCompare) = 1 Then
            TestHasSectionAnnotationIDPrefix = True
            Exit Function
        End If
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считает аннотации заданного типа; отсутствие модели не создает фиктивных записей.
Private Function TestCountSectionAnnotationType(ByVal model As CSectionModel, ByVal annotationType As String) As Long
    If model Is Nothing Then Exit Function

    Dim annotations As CSectionContours
    Set annotations = model.Contours

    Dim i As Long
    For i = 1 To annotations.Count
        If StrComp(annotations.SegmentType(i), annotationType, vbTextCompare) = 0 Then
            TestCountSectionAnnotationType = TestCountSectionAnnotationType + 1
        End If
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает индекс точной пары Type/ID либо ноль при отсутствии.
Private Function TestFindSectionAnnotationIndex(ByVal model As CSectionModel, ByVal annotationType As String, _
        ByVal annotationID As String) As Long
    If model Is Nothing Then Exit Function

    Dim annotations As CSectionAnnotations
    Set annotations = model.Annotations

    Dim i As Long
    For i = 1 To annotations.Count
        If StrComp(annotations.AnnotationType(i), annotationType, vbTextCompare) = 0 And _
                StrComp(annotations.AnnotationID(i), annotationID, vbTextCompare) = 0 Then
            TestFindSectionAnnotationIndex = i
            Exit Function
        End If
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Передает ошибочные дополнительные ряды только активной H1-грани тестового Rectangles.
' Остальные грани выключены, чтобы отказ относился к проверяемому набору.
Private Sub TestAssertRectanglesRebarRowsError(ByRef stats As TTestStats, ByVal name As String, ByVal h1Settings As Variant)
    On Error GoTo Expected
    Dim builder As CRebarRuleExecutor
    Set builder = New CRebarRuleExecutor
    Dim layout As CRebarLayout
    Set layout = TestFixtureAnnotatedLayout(TestFixtureRectanglesDefinition(Array(250#, 550#, 600#, 250#), _
        Array(h1Settings, TestEmptyFaceSettings(), TestEmptyFaceSettings(), TestEmptyFaceSettings())), builder)
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: " & name & "; expected validation error"
    Exit Sub
Expected:
    stats.Passed = stats.Passed + 1
    TestAppendLine stats, "OK: " & name
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считает созданные стержни по трассировочной части ID ряда.
Private Function TestCountBarsWithRow(ByVal layout As CRebarLayout, ByVal rowToken As String) As Long
    Dim i As Long
    For i = 1 To layout.Count
        If InStr(1, layout.BarID(i), rowToken, vbTextCompare) > 0 Then TestCountBarsWithRow = TestCountBarsWithRow + 1
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считает стержни по исходной грани/линии в ID раскладки.
Private Function TestCountBarsWithSource(ByVal layout As CRebarLayout, ByVal sourceToken As String) As Long
    Dim i As Long
    For i = 1 To layout.Count
        If InStr(1, layout.BarID(i), sourceToken, vbTextCompare) > 0 Then TestCountBarsWithSource = TestCountBarsWithSource + 1
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считает реальные стержни semantic-группы, а не число подписей этой группы.
Private Function TestCountBarsInAnnotationGroup(ByVal layout As CRebarLayout, ByVal groupName As String) As Long
    Dim i As Long
    For i = 1 To layout.Count
        If StrComp(layout.BarAnnotationGroupName(i), groupName, vbTextCompare) = 0 Then
            TestCountBarsInAnnotationGroup = TestCountBarsInAnnotationGroup + 1
        End If
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выбирает крайний Y только внутри нужной группы; первый стержень задает начальное значение.
Private Function TestExtremeYInAnnotationGroup(ByVal layout As CRebarLayout, ByVal groupName As String, ByVal findMax As Boolean) As Double
    Dim initialized As Boolean
    Dim i As Long
    For i = 1 To layout.Count
        If StrComp(layout.BarAnnotationGroupName(i), groupName, vbTextCompare) = 0 Then
            If Not initialized Then
                TestExtremeYInAnnotationGroup = layout.Y(i)
                initialized = True
            ElseIf findMax And layout.Y(i) > TestExtremeYInAnnotationGroup Then
                TestExtremeYInAnnotationGroup = layout.Y(i)
            ElseIf Not findMax And layout.Y(i) < TestExtremeYInAnnotationGroup Then
                TestExtremeYInAnnotationGroup = layout.Y(i)
            End If
        End If
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет Y каждого стержня горизонтальной группы и обязательное наличие стержней.
Private Sub TestAssertBarsInGroupHaveY(ByRef stats As TTestStats, ByVal layout As CRebarLayout, _
        ByVal groupName As String, ByVal expectedY As Double, ByVal prefix As String)
    Dim found As Boolean
    Dim i As Long
    For i = 1 To layout.Count
        If StrComp(layout.BarAnnotationGroupName(i), groupName, vbTextCompare) = 0 Then
            found = True
            TestAssertClose stats, prefix & ".barY." & CStr(i), layout.Y(i), expectedY, 0.000001
        End If
    Next i
    TestAssertTrue stats, prefix & ".hasBars", found
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет X каждого стержня вертикальной группы и обязательное наличие стержней.
Private Sub TestAssertBarsInGroupHaveX(ByRef stats As TTestStats, ByVal layout As CRebarLayout, _
        ByVal groupName As String, ByVal expectedX As Double, ByVal prefix As String)
    Dim found As Boolean
    Dim i As Long
    For i = 1 To layout.Count
        If StrComp(layout.BarAnnotationGroupName(i), groupName, vbTextCompare) = 0 Then
            found = True
            TestAssertClose stats, prefix & ".barX." & CStr(i), layout.X(i), expectedX, 0.000001
        End If
    Next i
    TestAssertTrue stats, prefix & ".hasBars", found
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выдает первый/последний индекс группы в порядке раскладки; оба нулевые при отсутствии.
Private Sub TestFindFirstLastBarInAnnotationGroup(ByVal layout As CRebarLayout, ByVal groupName As String, _
        ByRef firstIndex As Long, ByRef lastIndex As Long)
    Dim i As Long
    firstIndex = 0
    lastIndex = 0
    For i = 1 To layout.Count
        If StrComp(layout.BarAnnotationGroupName(i), groupName, vbTextCompare) = 0 Then
            If firstIndex = 0 Then firstIndex = i
            lastIndex = i
        End If
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит semantic-якорь подписи группы, не подменяя его координатами крайнего стержня.
Private Function TestFindRebarAnnotationAnchor(ByVal annotations As CSectionAnnotations, ByVal groupName As String) As Long
    Dim i As Long
    For i = 1 To annotations.Count
        If StrComp(annotations.AnnotationID(i), "REBAR_" & groupName, vbTextCompare) = 0 Then
            TestFindRebarAnnotationAnchor = i: Exit Function
        End If
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Assertions ниже сохраняют причины всех проверок группы без досрочного выхода suite.
Private Sub TestAssertTrue(ByRef stats As TTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        TestAppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        TestAppendLine stats, "FAIL: " & name
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает абсолютную разницу с заданным геометрическим допуском.
Private Sub TestAssertClose(ByRef stats As TTestStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal tolerance As Double)

    Dim diff As Double
    diff = Abs(actual - expected)
    If diff <= tolerance Then
        stats.Passed = stats.Passed + 1
        TestAppendLine stats, "OK: " & name & "; actual=" & TestFormatNumberInvariant(actual) & _
            "; expected=" & TestFormatNumberInvariant(expected) & "; absDiff=" & TestFormatNumberInvariant(diff)
    Else
        stats.Failed = stats.Failed + 1
        TestAppendLine stats, "FAIL: " & name & "; actual=" & TestFormatNumberInvariant(actual) & _
            "; expected=" & TestFormatNumberInvariant(expected) & "; absDiff=" & TestFormatNumberInvariant(diff)
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает относительную разницу; нулевой эталон проверяется без деления на него.
Private Sub TestAssertRelative(ByRef stats As TTestStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal relTolerance As Double)

    Dim relDiff As Double
    If Abs(expected) <= GEOM_TOLERANCE Then
        relDiff = Abs(actual - expected)
    Else
        relDiff = Abs((actual - expected) / expected)
    End If

    If relDiff <= relTolerance Then
        stats.Passed = stats.Passed + 1
        TestAppendLine stats, "OK: " & name & "; actual=" & TestFormatNumberInvariant(actual) & _
            "; expected=" & TestFormatNumberInvariant(expected) & "; relDiff=" & TestFormatNumberInvariant(relDiff)
    Else
        stats.Failed = stats.Failed + 1
        TestAppendLine stats, "FAIL: " & name & "; actual=" & TestFormatNumberInvariant(actual) & _
            "; expected=" & TestFormatNumberInvariant(expected) & "; relDiff=" & TestFormatNumberInvariant(relDiff)
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Helpers протокола сохраняют строки и форматируют числа с точкой для повторяемого сравнения.
Private Sub TestAppendLine(ByRef stats As TTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Форматирует число для текстового протокола с десятичной точкой; исходные расчётные значения не меняются.
Private Function TestFormatNumberInvariant(ByVal value As Double) As String
    TestFormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Проверяет, что повторный Initialize заменяет контур, валидность и причину,
' а запрос точки не принимает прежнюю геометрию после невалидного ввода.
Public Function RunAudit03GeometryLifecycleTests(Optional ByVal includeContourDetails As Boolean = False) As String
    On Error GoTo Failed
    Dim stats As TTestStats
    TestAudit03GeometryLifecycle stats, includeContourDetails
    TestAppendLine stats, "TOTAL_GEOMETRY_LIFECYCLE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03GeometryLifecycleTests = stats.Report
    Exit Function
Failed:
    RunAudit03GeometryLifecycleTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет раннюю защиту суммарного количества нескольких круговых рядов
' без попытки выделить гигантскую раскладку; возвращает отдельный протокол.
Public Function RunAudit03CircleCountGuardTests() As String
    On Error GoTo Failed
    Dim stats As TTestStats
    TestAudit03CircleCountGuard stats
    TestAppendLine stats, "TOTAL_CIRCLE_COUNT_GUARD: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03CircleCountGuardTests = stats.Report
    Exit Function
Failed:
    RunAudit03CircleCountGuardTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Малый радиус делает второй основной стержень почти совпадающим с первым:
' без ранней защиты старый маршрут быстро отказывает по координатам, а не
' запускает миллиардный цикл. Oracle требует именно причины количества.
' Затем тот же builder строит обычные ряды и сохраняет выключение первого.
Private Sub TestAudit03CircleCountGuard(ByRef stats As TTestStats)
    Dim builder As CRebarRuleExecutor, layout As CRebarLayout
    Set builder = New CRebarRuleExecutor
    Dim scenario As Long, count As Long, thirdDiameter As Double
    Dim errorNumber As Long, errorDescription As String, prefix As String
    For scenario = 2 To 3
        If scenario = 2 Then
            count = 1073741824
            thirdDiameter = 0#
        Else
            count = 715827883
            thirdDiameter = 0.001
        End If
        errorNumber = 0: errorDescription = vbNullString
        On Error Resume Next
        Set layout = TestFixtureAnnotatedLayout(TestFixtureCircleDefinition(Array(1#, 0.4, count, 0.001, 0.001, thirdDiameter)), builder)
        errorNumber = Err.Number: errorDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        prefix = "audit03.circle.countGuard." & CStr(scenario)
        TestAssertTrue stats, prefix & ".countReason", errorNumber = vbObjectError + 5530
        TestAssertTrue stats, prefix & ".table", InStr(1, errorDescription, "rngCircleGeometry", vbBinaryCompare) > 0
        TestAssertTrue stats, prefix & ".key", InStr(1, errorDescription, "количество", vbTextCompare) > 0
        TestAssertClose stats, prefix & ".beforeCoordinates", CDbl(builder.CreatedPositions), 0#, 0#
        TestAppendLine stats, "CIRCLE_COUNT_ERROR: rows=" & CStr(scenario) & "; number=" & CStr(errorNumber) & "; " & errorDescription
    Next scenario
    Set layout = TestFixtureMoveLayout(TestFixtureAnnotatedLayout(TestFixtureCircleDefinition(Array(1000#, 50#, 12, 16#, 16#)), builder), 0#, 0#)
    TestAssertTrue stats, "audit03.circle.countGuard.twoRowsRecovery", layout.Count = 24
    Set layout = TestFixtureMoveLayout(TestFixtureAnnotatedLayout(TestFixtureCircleDefinition(Array(1000#, 50#, 12, 16#, 16#, 12#)), builder), 0#, 0#)
    TestAssertTrue stats, "audit03.circle.countGuard.threeRowsRecovery", layout.Count = 36
    Set layout = TestFixtureMoveLayout(TestFixtureAnnotatedLayout(TestFixtureCircleDefinition(Array(1000#, 50#, 1073741824, 0#, 16#, 12#)), builder), 0#, 0#)
    TestAssertTrue stats, "audit03.circle.countGuard.disabledBase", layout.Count = 0
    TestAssertClose stats, "audit03.circle.countGuard.disabledRadius", CDbl(builder.CreatedPositions), 0#, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет адресный контекст невозможного дополнительного ряда через
' настоящий reader и временно перемещенный rngCircleGeometry. Геометрический
' код отказа сохраняется; после ошибки тот же builder строит допустимый ряд.
Public Function RunAudit03CirclePlacementAddressTests() As String
    Dim stats As TTestStats
    TestAudit03CirclePlacementAddresses stats
    TestAppendLine stats, "TOTAL_CIRCLE_PLACEMENT_ADDRESSES: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03CirclePlacementAddressTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет только собственную временную таблицу. Исходный Name, активный лист
' и режим предупреждений восстанавливаются при успехе и при runtime-ошибке.
Private Sub TestAudit03CirclePlacementAddresses(ByRef stats As TTestStats)
    Dim sheet As Object, previousSheet As Object, geometryRef As String
    Dim geometry As Object, rebars As Object, targetGeometry As Object, targetRebars As Object
    Dim config As Object, saved As Variant, alerts As Boolean, position As Variant, extraRow As Long, location As Variant
    Dim units As CUnitSystem, builder As CRebarRuleExecutor, layout As CRebarLayout
    Dim prefix As String, number As Long, reason As String, axis As Object, diameter As Object
    On Error GoTo FailedRun
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    geometryRef = ThisWorkbook.Names("rngCircleGeometry").RefersTo
    modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
    Set geometry = ThisWorkbook.Names("rngCircleGeometry").RefersToRange
    Set previousSheet = ThisWorkbook.ActiveSheet: alerts = Application.DisplayAlerts
    Set sheet = ThisWorkbook.Worksheets.Add
    Set units = New CUnitSystem: units.InitializeDefaults
    Set builder = New CRebarRuleExecutor
    For Each position In Array("E10", "BH800")
        Set targetGeometry = sheet.Range(CStr(position)).Resize(geometry.Rows.Count, geometry.Columns.Count)
        ' Проверяем перенос адресов; ключи фиксируем уже вычисленным снимком,
        ' чтобы относительная формула Extra не ссылалась на старое место Main.
        targetGeometry.Value2 = geometry.Value2
        ThisWorkbook.Names("rngCircleGeometry").RefersTo = "='" & sheet.Name & "'!" & targetGeometry.Address
        TestFixtureCircleFieldCell(ThisWorkbook, "Circle.Diameter").Value2 = 300#
        Set axis = TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.AxisDistance"): axis.Value2 = 20#
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Diameter").Value2 = 20#
        TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Count").Value2 = 8
        For extraRow = 2 To 3
            For Each location In Array("Stacked", "SideBySide")
                TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Diameter2").Value2 = 10#
                TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Diameter3").Value2 = 0#
                TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Loc2row").Value2 = CStr(location)
                TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Loc3row").Value2 = CStr(location)
                Set diameter = TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Diameter" & CStr(extraRow))
                diameter.Value2 = 1000#: Set layout = Nothing: number = 0: reason = ""
                On Error Resume Next
                Set layout = TestFixtureAnnotatedLayout(TestFixtureWorkbookDefinition("Circle", ThisWorkbook, units), builder)
                number = Err.Number: reason = Err.Description: Err.Clear
                On Error GoTo FailedRun
                prefix = "audit03.circle.placementAddress." & CStr(position) & "." & CStr(extraRow) & "." & CStr(location)
                TestAssertTrue stats, prefix & ".sameCode", number = vbObjectError + 5530
                TestAssertTrue stats, prefix & ".noLayout", layout Is Nothing
                TestAssertTrue stats, prefix & ".sheet", InStr(1, reason, sheet.Name, vbBinaryCompare) > 0
                TestAssertTrue stats, prefix & ".activeDiameter", InStr(1, reason, diameter.Address(False, False), vbBinaryCompare) > 0
                TestAssertTrue stats, prefix & ".rowCaption", InStr(1, reason, "Круг:" & vbCrLf & "Дополнительный ряд арматуры " & CStr(extraRow), vbBinaryCompare) > 0
                TestAssertTrue stats, prefix & ".axisUnchanged", axis.Value2 = 20#
                TestAssertTrue stats, prefix & ".action", InStr(1, reason, "уменьшите", vbTextCompare) > 0
                TestAppendLine stats, "CIRCLE_PLACEMENT_ERROR: " & prefix & "; " & reason
                diameter.Value2 = 10#
                Set layout = TestFixtureAnnotatedLayout(TestFixtureWorkbookDefinition("Circle", ThisWorkbook, units), builder)
                TestAssertTrue stats, prefix & ".recovery", layout.Count = 8 * extraRow
            Next location
        Next extraRow
    Next position
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: audit03.circle.placementAddress.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    On Error Resume Next
    If Len(geometryRef) > 0 Then ThisWorkbook.Names("rngCircleGeometry").RefersTo = geometryRef
    If Not config Is Nothing And IsArray(saved) Then TestFixtureRestoreInputs config, saved
    If Not previousSheet Is Nothing Then previousSheet.Activate
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = alerts
    On Error GoTo 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' невалидное описание не должно превращаться в пригодную область
' или неявно возвращать прежний снимок после повторной инициализации формы.
Private Function TestGeometryPreparationFails(ByVal definition As CSectionDefinition) As Boolean
    Dim geometry As CGeometryRegion, number As Long, description As String
    On Error Resume Next
    Set geometry = TestFixtureGeometry(definition.ConcreteDomain)
    number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    TestGeometryPreparationFails = number <> 0 And Len(description) > 0 And geometry Is Nothing
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает подготовленные области после valid-invalid-valid переходов.
Private Sub TestAudit03GeometryLifecycle(ByRef stats As TTestStats, Optional ByVal includeContourDetails As Boolean = False)
    Dim rounded As CSectionDefinition: Set rounded = New CSectionDefinition
    Dim hollow As CSectionDefinition: Set hollow = New CSectionDefinition
    Dim message As String, repeatedMessage As String
    TestAssertTrue stats, "audit03.geometry.rounded.uninitialized", Not TestPublishedDefinition(rounded)
    TestAssertTrue stats, "audit03.geometry.rounded.uninitializedPreparation", TestGeometryPreparationFails(rounded)
    TestAssertTrue stats, "audit03.geometry.hollow.uninitialized", Not TestPublishedDefinition(hollow)
    TestAssertTrue stats, "audit03.geometry.hollow.uninitializedPreparation", TestGeometryPreparationFails(hollow)

    Set rounded = TestFixtureRoundedDefinition(Array(300#, 200#, 0#, 0#, 20#, 20#, 0#, 0#))
    TestAssertTrue stats, "audit03.geometry.rounded.firstValid", TestPublishedDefinition(rounded)
    TestAssertTrue stats, "audit03.geometry.rounded.firstCenter", TestFixtureGeometry(rounded.ConcreteDomain).ContainsPoint(0#, 0#)
    Set rounded = TestFailedGeometryDefinition(False, message)
    TestAssertTrue stats, "audit03.geometry.rounded.invalid", Not TestPublishedDefinition(rounded)
    TestAssertTrue stats, "audit03.geometry.rounded.invalidPreparation", TestGeometryPreparationFails(rounded)
    Set rounded = TestFailedGeometryDefinition(False, repeatedMessage)
    TestAssertTrue stats, "audit03.geometry.rounded.repeatInvalid", Not TestPublishedDefinition(rounded)
    TestAssertTrue stats, "audit03.geometry.rounded.stableReason", message = repeatedMessage And Len(message) > 0
    Dim roundedParameters As Variant
    roundedParameters = Array(400#, 180#, 0#, 70#, 20#, 25#, 0#, 10#)
    Set rounded = TestFixtureRoundedDefinition(roundedParameters)
    Dim freshRounded As CSectionDefinition
    Dim freshRoundedParameters As Variant
    freshRoundedParameters = Array(400#, 180#, 0#, 70#, 20#, 25#, 0#, 10#)
    Set freshRounded = TestFixtureRoundedDefinition(freshRoundedParameters)
    TestAssertTrue stats, "audit03.geometry.rounded.revalidated", TestPublishedDefinition(rounded)
    If includeContourDetails Then
        Dim contourX() As Double, contourY() As Double, contourIndex As Long
        TestFixtureMoveGeometry(rounded, 50#, -30#).GetExtremePoints contourX, contourY
        For contourIndex = LBound(contourX) To UBound(contourX)
            TestAppendLine stats, "CONTOUR: Rounded; i=" & CStr(contourIndex) & "; x=" & CStr(contourX(contourIndex)) & _
                "; y=" & CStr(contourY(contourIndex)) & "; deltaFromBottom=" & CStr((contourY(contourIndex) + 120#) * 1000000000000#)
        Next contourIndex
    End If

    Dim hollowParameters As Variant
    hollowParameters = Array(500#, 800#, 30#, 200#, 500#, 20#, 40#, -30#)
    Set hollow = TestFixtureHollowDefinition(hollowParameters)
    TestAssertTrue stats, "audit03.geometry.hollow.firstValid", TestPublishedDefinition(hollow)
    TestAssertTrue stats, "audit03.geometry.hollow.opening", Not TestFixtureGeometry(hollow.ConcreteDomain).ContainsPoint(40#, -30#)
    hollowParameters = Array(500#, 800#, 30#, 600#, 500#, 20#, 40#, -30#)
    Set hollow = TestFailedGeometryDefinition(True, message)
    TestAssertTrue stats, "audit03.geometry.hollow.invalid", Not TestPublishedDefinition(hollow)
    TestAssertTrue stats, "audit03.geometry.hollow.invalidPreparation", TestGeometryPreparationFails(hollow)
    Set hollow = TestFailedGeometryDefinition(True, repeatedMessage)
    TestAssertTrue stats, "audit03.geometry.hollow.repeatInvalid", Not TestPublishedDefinition(hollow)
    TestAssertTrue stats, "audit03.geometry.hollow.stableReason", message = repeatedMessage And Len(message) > 0
    hollowParameters = Array(450#, 650#, 40#, 180#, 300#, 20#, -30#, 20#)
    Set hollow = TestFixtureHollowDefinition(hollowParameters)
    Dim freshHollow As CSectionDefinition
    Dim freshHollowParameters As Variant
    freshHollowParameters = Array(450#, 650#, 40#, 180#, 300#, 20#, -30#, 20#)
    Set freshHollow = TestFixtureHollowDefinition(freshHollowParameters)
    TestAssertTrue stats, "audit03.geometry.hollow.revalidated", TestPublishedDefinition(hollow)
    TestAssertTrue stats, "audit03.geometry.hollow.newOpening", Not TestFixtureGeometry(hollow.ConcreteDomain).ContainsPoint(-30#, 20#)

    Dim ix As Long, iy As Long
    Dim roundedMismatch As Long, hollowMismatch As Long
    Dim roundedRow As String, hollowRow As String
    Dim roundedDomain As CGeometryRegion, freshRoundedDomain As CGeometryRegion
    Dim hollowDomain As CGeometryRegion, freshHollowDomain As CGeometryRegion
    Set roundedDomain = TestFixtureMoveGeometry(rounded, 50#, -30#): Set freshRoundedDomain = TestFixtureMoveGeometry(freshRounded, 50#, -30#)
    Set hollowDomain = TestFixtureGeometry(hollow.ConcreteDomain): Set freshHollowDomain = TestFixtureGeometry(freshHollow.ConcreteDomain)
    For iy = -18 To 18
        roundedRow = vbNullString
        hollowRow = vbNullString
        For ix = -24 To 24
            If roundedDomain.ContainsPoint(ix * 15#, iy * 20#) <> freshRoundedDomain.ContainsPoint(ix * 15#, iy * 20#) Then roundedMismatch = roundedMismatch + 1
            If hollowDomain.ContainsPoint(ix * 15#, iy * 20#) <> freshHollowDomain.ContainsPoint(ix * 15#, iy * 20#) Then hollowMismatch = hollowMismatch + 1
            If roundedDomain.ContainsPoint(ix * 15#, iy * 20#) Then roundedRow = roundedRow & "1" Else roundedRow = roundedRow & "0"
            If hollowDomain.ContainsPoint(ix * 15#, iy * 20#) Then hollowRow = hollowRow & "1" Else hollowRow = hollowRow & "0"
        Next ix
        TestAppendLine stats, "GRID: Rounded; y=" & CStr(iy * 20#) & "; " & roundedRow
        TestAppendLine stats, "GRID: Hollow; y=" & CStr(iy * 20#) & "; " & hollowRow
    Next iy
    TestAssertTrue stats, "audit03.geometry.rounded.reinitGrid", roundedMismatch = 0
    TestAssertTrue stats, "audit03.geometry.hollow.reinitGrid", hollowMismatch = 0
    TestAssertTrue stats, "audit03.geometry.hollow.outerBoundary", TestFixtureGeometry(hollow.ConcreteDomain).ContainsPoint(225#, 0#)
    TestAssertTrue stats, "audit03.geometry.hollow.openingBoundary", Not TestFixtureGeometry(hollow.ConcreteDomain).ContainsPoint(60#, 20#)
    Dim available As Boolean, freshAvailable As Boolean
    TestAssertClose stats, "audit03.geometry.rounded.reinitArea", TestFixtureGeometry(rounded.ConcreteDomain).AnalyticalArea(available), TestFixtureGeometry(freshRounded.ConcreteDomain).AnalyticalArea(freshAvailable), 0#
    TestAssertTrue stats, "audit03.geometry.rounded.areaAvailable", available And freshAvailable
    TestAssertClose stats, "audit03.geometry.hollow.reinitArea", TestFixtureGeometry(hollow.ConcreteDomain).AnalyticalArea(available), TestFixtureGeometry(freshHollow.ConcreteDomain).AnalyticalArea(freshAvailable), 0#
    TestAssertTrue stats, "audit03.geometry.hollow.areaAvailable", available And freshAvailable
    TestAppendLine stats, "CONST: GEOM_PI deltaFrom3x1e15=" & CStr((GEOM_PI - 3#) * 1E+15)
    TestAssertClose stats, "audit03.geometry.piFullPrecision", GEOM_PI, 4# * Atn(1#), 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет проем целиком внутри базовой ячейки: все ее углы находятся
' в бетоне, но это не разрешает включать пустоту в площадь сетки.
' Границы прямоугольных проемов совпадают с границами подъячеек 10x10 мм,
' поэтому независимый oracle площади не требует погрешности аппроксимации.
Private Sub TestAudit03HollowContainedOpeningMesh(ByRef stats As TTestStats)
    Dim hollow As CGeometryRegion, mesh As CFiberMeshBuilder
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, definition As CSectionDefinition
    Dim scenario As Long, i As Long, area As Double, invalidCenters As Long
    Dim openingWidth As Double, openingHeight As Double, offsetX As Double, offsetY As Double
    Dim sx As Double, sy As Double, ix As Double, iy As Double, ixy As Double, fiberArea As Double
    Dim expectedArea As Double, openingArea As Double, cx As Double, cy As Double
    Dim prefix As String
    Set mesh = New CFiberMeshBuilder
    For scenario = 0 To 2
        openingWidth = 40#: openingHeight = 40#: offsetX = 50#: offsetY = 50#
        If scenario = 1 Then
            openingWidth = 20#: openingHeight = 40#: offsetX = 70#: offsetY = 30#
        ElseIf scenario = 2 Then
            offsetX = 0#: offsetY = 0#
        End If
        Set facade = New CSectionDefinitionBuilder: facade.Initialize "ContainedVoidFixture"
        Set path = facade.RectanglePath("Boundary", -200#, -200#, 400#, 400#): facade.Outer "OUTER", path.ID
        Set path = facade.RectanglePath("Void", offsetX - openingWidth / 2#, offsetY - openingHeight / 2#, openingWidth, openingHeight)
        facade.Opening "VOID", path.ID, "OUTER"
        Set definition = facade.Finish: Set hollow = TestFixtureGeometry(definition.ConcreteDomain)
        mesh.BuildMesh hollow, 100#, 100#, 2, 10
        area = 0#: invalidCenters = 0#: sx = 0#: sy = 0#: ix = 0#: iy = 0#: ixy = 0#
        For i = 1 To mesh.FiberCount
            fiberArea = mesh.FiberArea(i)
            area = area + fiberArea
            sx = sx + fiberArea * mesh.FiberX(i): sy = sy + fiberArea * mesh.FiberY(i)
            ' Суммируем целочисленный числитель и делим один раз: локальная
            ' инерция подъячеек не накапливает округление повторного деления на 12.
            ix = ix + fiberArea * (12# * mesh.FiberY(i) ^ 2 + mesh.FiberHeight(i) ^ 2)
            iy = iy + fiberArea * (12# * mesh.FiberX(i) ^ 2 + mesh.FiberWidth(i) ^ 2)
            ixy = ixy + fiberArea * mesh.FiberX(i) * mesh.FiberY(i)
            If Not hollow.ContainsPoint(mesh.FiberX(i), mesh.FiberY(i)) Then invalidCenters = invalidCenters + 1
        Next i
        prefix = "audit03.hollow.containedOpening." & CStr(scenario)
        openingArea = openingWidth * openingHeight: expectedArea = 160000# - openingArea
        cx = -openingArea * offsetX / expectedArea: cy = -openingArea * offsetY / expectedArea
        TestAssertClose stats, prefix & ".area", area, expectedArea, 0.000001
        TestAssertClose stats, prefix & ".centroidX", sx / area, cx, 0.000000001
        TestAssertClose stats, prefix & ".centroidY", sy / area, cy, 0.000000001
        TestAssertClose stats, prefix & ".Ix", ix / 12# - sy * sy / area, _
            400# ^ 4 / 12# - openingArea * (openingHeight ^ 2 / 12# + offsetY ^ 2) - expectedArea * cy ^ 2, 0.000001
        TestAssertClose stats, prefix & ".Iy", iy / 12# - sx * sx / area, _
            400# ^ 4 / 12# - openingArea * (openingWidth ^ 2 / 12# + offsetX ^ 2) - expectedArea * cx ^ 2, 0.000001
        TestAssertClose stats, prefix & ".Ixy", ixy - sx * sy / area, _
            -openingArea * offsetX * offsetY - expectedArea * cx * cy, 0.000001
        TestAssertTrue stats, prefix & ".allCentersInConcrete", invalidCenters = 0
        TestAppendLine stats, "HOLLOW_CELL_MESH: " & prefix & "; fibers=" & CStr(mesh.FiberCount) & _
            "; area=" & CStr(area) & "; invalidCenters=" & CStr(invalidCenters)
    Next scenario
    Dim rectangle As CGeometryRegion
    Set rectangle = TestFixtureBox(200#, 200#)
    mesh.BuildMesh rectangle, 100#, 100#, 2, 10
    Dim fullCellCount As Long
    For i = 1 To mesh.FiberCount
        If mesh.FiberWidth(i) = 100# And mesh.FiberHeight(i) = 100# Then fullCellCount = fullCellCount + 1
    Next i
    TestAssertTrue stats, "audit03.hollow.containedOpening.convexFastPath", fullCellCount > 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает отдельный протокол проемов внутри грубой ячейки и сохранения
' быстрого пути сплошного прямоугольника; исходная геометрия книги не читается.
Public Function RunAudit03HollowContainedOpeningMeshTests() As String
    On Error GoTo Failed
    Dim stats As TTestStats
    TestAudit03HollowContainedOpeningMesh stats
    TestAppendLine stats, "TOTAL_HOLLOW_CELL_MESH: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03HollowContainedOpeningMeshTests = stats.Report
    Exit Function
Failed:
    RunAudit03HollowContainedOpeningMeshTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function


















' -------------------- ДЛЯ ТЕСТОВ --------------------
' готовое описание сечения проверяется общим способом вместо прежнего изменяемого
' свойства формы IsValid.
Private Function TestPublishedDefinition(ByVal definition As CSectionDefinition) As Boolean
    If Not definition Is Nothing Then TestPublishedDefinition = definition.IsSealed
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Получает ожидаемую причину отказа для слишком большого отверстия либо недопустимого скругления, не подменяя её своей ошибкой.
Private Function TestFailedGeometryDefinition(ByVal hollow As Boolean, ByRef reason As String) As CSectionDefinition
    On Error GoTo Failed
    If hollow Then
        Set TestFailedGeometryDefinition = TestFixtureHollowDefinition(Array(500#, 800#, 30#, 600#, 500#, 20#, 40#, -30#))
    Else
        Set TestFailedGeometryDefinition = TestFixtureRoundedDefinition(Array(300#, 200#, 0#, 0#, 1000#, 1000#, 0#, 0#))
    End If
    Exit Function
Failed:
    reason = Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Короткий прогон исправленного размерного guard, штатных прямоугольников СП 35 и безопасного состояния геометрии.
Public Function TestRunWorkingGeometryReviewTests() As String
    On Error GoTo Failed
    Dim stats As TTestStats
    TestGeometryReviewCoreScale stats
    TestGeometryReviewSP35Rectangles stats
    TestGeometryReviewTranslationInertia stats
    TestGeometryReviewMeshFailureState stats
    TestGeometryReviewConcreteAddFailureState stats
    TestSectionModelFromGeneratedGeometry stats
    TestGeometryReviewLargeAngles stats
    TestAppendLine stats, "TOTAL_WORKING_GEOMETRY_RELIABILITY: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    TestRunWorkingGeometryReviewTests = stats.Report
    Exit Function
Failed:
    TestRunWorkingGeometryReviewTests = stats.Report & "FAIL: workingGeometry.runtime=" & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ядро прямоугольника проверяется независимо по аналитическому ромбу при пяти масштабах, трёх поворотах и пяти направлениях.
Private Sub TestGeometryReviewCoreScale(ByRef stats As TTestStats)
    Dim scaleFactor As Variant, angle As Variant, direction As Variant
    For Each scaleFactor In Array(0.1, 1#, 2#, 3#, 10#)
        For Each angle In Array(0#, GEOM_PI / 7#, GEOM_PI / 2#)
            For Each direction In Array(0#, GEOM_PI / 5#, GEOM_PI / 2#, GEOM_PI, 7# * GEOM_PI / 5#)
                TestGeometryReviewCoreRectangle stats, 300# * CDbl(scaleFactor), 1000# * CDbl(scaleFactor), CDbl(angle), CDbl(direction)
            Next direction
        Next angle
    Next scaleFactor
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' В локальных осях прямоугольника расстояние до ядра равно 1/[6*(|nx|/B+|ny|/H)], без обращения к рабочему guard.
Private Sub TestGeometryReviewCoreRectangle(ByRef stats As TTestStats, ByVal width As Double, ByVal height As Double, _
        ByVal rotation As Double, ByVal direction As Double)
    Dim model As CSectionModel, props As CSectionPropertiesCalculator, ix As Double, iy As Double
    Dim globalIx As Double, globalIy As Double, globalIxy As Double, c As Double, sn As Double, expected As Double
    c = Cos(rotation): sn = Sin(rotation)
    ix = width * height ^ 3 / 12#: iy = height * width ^ 3 / 12#
    globalIx = ix * c * c + iy * sn * sn: globalIy = ix * sn * sn + iy * c * c
    globalIxy = (iy - ix) * sn * c
    Set model = New CSectionModel
    model.AddConcreteElement 0#, 0#, width * height, 1, shapeType:="Rectangle", width:=width, height:=height, rotation:=rotation, localIx:=globalIx, localIy:=globalIy, localIxy:=globalIxy
    Set props = New CSectionPropertiesCalculator: props.CalculateConcrete model
    expected = 1# / (6# * (Abs(Cos(direction - rotation)) / width + Abs(Sin(direction - rotation)) / height))
    TestAssertClose stats, "review.core.B=" & CStr(width) & ".H=" & CStr(height) & ".rot=" & CStr(rotation) & ".dir=" & CStr(direction), _
        props.CoreDistanceAlong(model, Cos(direction), Sin(direction), False), expected, expected * 0.000000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Прямоугольник 200x600 уже принят в проверке главных осей СП 35; его подобный 2000x6000 должен давать в десять раз большие r.
Private Sub TestGeometryReviewSP35Rectangles(ByRef stats As TTestStats)
    Dim scaleFactor As Variant, model As CSectionModel, props As CSectionPropertiesCalculator, width As Double, height As Double
    For Each scaleFactor In Array(1#, 10#)
        width = 200# * CDbl(scaleFactor): height = 600# * CDbl(scaleFactor)
        Set model = New CSectionModel
        model.AddConcreteElement 0#, 0#, width * height, 1, shapeType:="Rectangle", width:=width, height:=height
        Set props = New CSectionPropertiesCalculator: props.CalculateConcrete model
        TestAssertClose stats, "review.sp35.rectangle.rX." & CStr(scaleFactor), props.CoreDistanceAlong(model, 1#, 0#, False), width / 6#, 0.000001
        TestAssertClose stats, "review.sp35.rectangle.rY." & CStr(scaleFactor), props.CoreDistanceAlong(model, 0#, 1#, False), height / 6#, 0.000001
        TestAssertClose stats, "review.sp35.rectangle.plane1." & CStr(scaleFactor), props.PrincipalPlaneCoreDistance(model, 1, True, False), height / 6#, 0.000001
        TestAssertClose stats, "review.sp35.rectangle.plane2." & CStr(scaleFactor), props.PrincipalPlaneCoreDistance(model, 2, True, False), width / 6#, 0.000001
    Next scaleFactor
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ошибка изменяемой повторной сборки сбрасывает старый результат, а отказ перестройки Frozen-сетки сохраняет опубликованные элементы.
Private Sub TestGeometryReviewMeshFailureState(ByRef stats As TTestStats)
    Dim geometry As CGeometryRegion, mesh As CFiberMeshBuilder, oldCount As Long, failed As Long
    Dim builder As CSectionModelBuilder, model As CSectionModel
    Set geometry = TestFixtureDisk(20#)
    Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geometry, 2#, 2#
    On Error Resume Next
    mesh.BuildMesh geometry, 0#, 2#: failed = Err.Number: Err.Clear
    On Error GoTo 0
    TestAssertTrue stats, "review.mesh.invalidRejected", failed <> 0
    TestAssertClose stats, "review.mesh.noStaleFibers", mesh.FiberCount, 0#, 0#
    TestAssertClose stats, "review.mesh.noStaleSeconds", mesh.TestBuildSeconds, 0#, 0#
    Set builder = New CSectionModelBuilder
    On Error Resume Next
    Set model = builder.BuildFromGenerated(mesh, Nothing): failed = Err.Number: Err.Clear
    On Error GoTo 0
    TestAssertTrue stats, "review.mesh.failedNotPublished", failed <> 0
    mesh.BuildMesh geometry, 2#, 2#: oldCount = mesh.FiberCount: mesh.Freeze
    On Error Resume Next
    mesh.BuildMesh geometry, 0#, 2#: failed = Err.Number: Err.Clear
    On Error GoTo 0
    TestAssertTrue stats, "review.mesh.frozenRebuildRejected", failed <> 0
    TestAssertClose stats, "review.mesh.frozenCountPreserved", mesh.FiberCount, oldCount, 0#
    TestAssertTrue stats, "review.mesh.frozenReadOnly", mesh.IsReadOnly
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отказ добавления не публикует незаконченный элемент; следующее успешное добавление корректно меняет версию и инвалидирует все подготовленные свойства.
Private Sub TestGeometryReviewConcreteAddFailureState(ByRef stats As TTestStats)
    Dim model As CSectionModel, control As CSectionModel, props As CSectionPropertiesCalculator
    Dim expected As CSectionPropertiesCalculator, snapshot As CSectionPropertiesCalculator
    Dim mode As Long, failure As Long, oldRevision As Long, oldBuildCount As Long, prefix As String
    Dim minValue As Double, maxValue As Double, cx As Double, cy As Double, area As Double, hiddenID As String
    Set model = New CSectionModel: Set control = New CSectionModel
    model.AddConcreteElement 5#, -7#, 600#, width:=30#, height:=20#
    control.AddConcreteElement 5#, -7#, 600#, width:=30#, height:=20#
    Set props = New CSectionPropertiesCalculator: props.EnsureConcrete model
    props.CalculateProjection model, 1#, 0#, False, minValue, maxValue
    model.ConcreteMeshCentroid cx, cy, area
    area = model.MaximumElementRadiusFromConcreteCenter ' Заранее заполняем отдельный кэш радиуса, чтобы отказ не скрывался за первым чтением после него.
    Set snapshot = New CSectionPropertiesCalculator: snapshot.PrepareSnapshot model
    oldRevision = model.Revision: oldBuildCount = props.TestPreparationBuildCount
    For mode = 1 To 2
        On Error Resume Next
        If mode = 1 Then
            model.AddConcreteElement 100#, 100#, 1E+200, width:=1E+100, height:=1E+100
        Else
            model.AddConcreteElement 100#, 100#, 600#, width:=30#, height:=20#, localIx:=1E+200, localIy:=1E+200
        End If
        failure = Err.Number: Err.Clear
        On Error GoTo 0
        prefix = "review.addFailure.mode=" & CStr(mode)
        TestAssertTrue stats, prefix & ".overflow", failure = 6
        TestAssertClose stats, prefix & ".countUnchanged", model.ConcreteCount, 1#, 0#
        TestAssertClose stats, prefix & ".revisionUnchanged", model.Revision, oldRevision, 0#
        props.EnsureConcrete model
        TestAssertClose stats, prefix & ".preparedContextReusable", props.TestPreparationBuildCount, oldBuildCount, 0#
        TestAssertClose stats, prefix & ".area", props.Area, 600#, 0#
        TestAssertClose stats, prefix & ".Ix", props.Ixc, 20000#, 0#
        TestAssertClose stats, prefix & ".Iy", props.Iyc, 45000#, 0#
        TestAssertClose stats, prefix & ".Ixy", props.Ixyc, 0#, 0#
        props.CalculateProjection model, 1#, 0#, False, minValue, maxValue
        TestAssertClose stats, prefix & ".cachedProjectionMin", minValue, -10#, 0#
        TestAssertClose stats, prefix & ".cachedProjectionMax", maxValue, 20#, 0#
        model.ConcreteMeshCentroid cx, cy, area
        TestAssertClose stats, prefix & ".cachedCenterX", cx, 5#, 0#
        TestAssertClose stats, prefix & ".cachedCenterY", cy, -7#, 0#
        TestAssertClose stats, prefix & ".cachedCenterArea", area, 600#, 0#
        TestAssertClose stats, prefix & ".cachedRadius", model.MaximumElementRadiusFromConcreteCenter, 0#, 0#
        On Error Resume Next
        snapshot.ValidateSnapshot model: failure = Err.Number: Err.Clear
        On Error GoTo 0
        TestAssertTrue stats, prefix & ".snapshotStillValid", failure = 0
        On Error Resume Next
        hiddenID = model.ConcreteID(2): failure = Err.Number: Err.Clear
        On Error GoTo 0
        TestAssertTrue stats, prefix & ".partialSlotInvisible", failure <> 0
        TestAssertTrue stats, prefix & ".originalElement", model.ConcreteID(1) = "C1" And model.ConcreteX(1) = 5# And model.ConcreteY(1) = -7#
        TestAssertClose stats, prefix & ".queriesDoNotChangeRevision", model.Revision, oldRevision, 0#
    Next mode
    model.AddConcreteElement 35#, 13#, 300#, 2, "Recovered", "H2", width:=15#, height:=20#, comment:="second valid"
    control.AddConcreteElement 35#, 13#, 300#, 2, "Recovered", "H2", width:=15#, height:=20#, comment:="second valid"
    TestAssertClose stats, "review.addRecovery.count", model.ConcreteCount, 2#, 0#
    TestAssertClose stats, "review.addRecovery.revision", model.Revision, oldRevision + 1, 0#
    TestAssertTrue stats, "review.addRecovery.noIdGap", model.ConcreteID(2) = "C2"
    TestAssertTrue stats, "review.addRecovery.noPartialComment", model.ConcreteComment(2) = "second valid"
    props.EnsureConcrete model
    TestAssertClose stats, "review.addRecovery.contextRebuiltOnce", props.TestPreparationBuildCount, oldBuildCount + 1, 0#
    Set expected = New CSectionPropertiesCalculator: expected.CalculateConcrete control
    TestAssertClose stats, "review.addRecovery.area", props.Area, expected.Area, 0#
    TestAssertClose stats, "review.addRecovery.centroidX", props.CentroidX, expected.CentroidX, 0#
    TestAssertClose stats, "review.addRecovery.centroidY", props.CentroidY, expected.CentroidY, 0#
    TestAssertClose stats, "review.addRecovery.Ix", props.Ixc, expected.Ixc, 0#
    TestAssertClose stats, "review.addRecovery.Iy", props.Iyc, expected.Iyc, 0#
    TestAssertClose stats, "review.addRecovery.Ixy", props.Ixyc, expected.Ixyc, 0#
    TestAssertClose stats, "review.addRecovery.principalI1", props.PrincipalI1, expected.PrincipalI1, 0#
    TestAssertClose stats, "review.addRecovery.principalI2", props.PrincipalI2, expected.PrincipalI2, 0#
    props.CalculateProjection model, 1#, 0#, False, minValue, maxValue
    TestAssertClose stats, "review.addRecovery.projectionMin", minValue, -10#, 0#
    TestAssertClose stats, "review.addRecovery.projectionMax", maxValue, 42.5, 0#
    model.ConcreteMeshCentroid cx, cy, area
    TestAssertClose stats, "review.addRecovery.centerX", cx, expected.CentroidX, 0#
    TestAssertClose stats, "review.addRecovery.centerY", cy, expected.CentroidY, 0#
    TestAssertClose stats, "review.addRecovery.centerArea", area, expected.Area, 0#
    TestAssertClose stats, "review.addRecovery.radiusRebuilt", model.MaximumElementRadiusFromConcreteCenter, control.MaximumElementRadiusFromConcreteCenter, 0#
    On Error Resume Next
    snapshot.ValidateSnapshot model: failure = Err.Number: Err.Clear
    On Error GoTo 0
    TestAssertTrue stats, "review.addRecovery.oldSnapshotRejected", failure <> 0
    TestAssertTrue stats, "review.addRecovery.originalElement", model.ConcreteID(1) = "C1" And model.ConcreteX(1) = 5# And model.ConcreteY(1) = -7#
    TestAssertTrue stats, "review.addRecovery.replacedSlot", model.ConcreteX(2) = 35# And model.ConcreteY(2) = 13# And model.ConcreteMaterialID(2) = 2 And model.ConcreteSourceName(2) = "Recovered" And model.ConcreteSourceHandle(2) = "H2"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' В штатном диапазоне проверяется точное равенство прежнему нормализатору; огромные углы проверяются только на исправленной реализации.
Private Sub TestGeometryReviewLargeAngles(ByRef stats As TTestStats)
    Dim angle As Variant, model As CSectionModel, expected As Double, actual As Double
    For Each angle In Array(-2# * GEOM_PI, -GEOM_PI, -GEOM_PI / 2#, -0.1, 0#, 0.1, GEOM_PI / 2#, GEOM_PI, 2# * GEOM_PI)
        Set model = New CSectionModel
        model.AddConcreteElement 0#, 0#, 200#, 1, shapeType:="Rectangle", width:=20#, height:=10#, rotation:=CDbl(angle)
        expected = TestGeometryReviewLegacyAngle(CDbl(angle))
        TestAssertTrue stats, "review.angle.boundedExact." & CStr(angle), model.ConcreteRotation(1) = expected
    Next angle
    For Each angle In Array(-1000000#, 1000000#, -1000# * GEOM_PI - 0.23, 1000# * GEOM_PI + 0.23)
        Set model = New CSectionModel
        model.AddConcreteElement 0#, 0#, 200#, 1, shapeType:="Rectangle", width:=20#, height:=10#, rotation:=CDbl(angle)
        actual = model.ConcreteRotation(1)
        TestAssertTrue stats, "review.angle.range." & CStr(angle), actual > -GEOM_PI / 2# And actual <= GEOM_PI / 2#
        TestAssertClose stats, "review.angle.axisEquivalent." & CStr(angle), Cos(actual) * Sin(CDbl(angle)) - Sin(actual) * Cos(CDbl(angle)), 0#, 0.00000001
        expected = (20# * Abs(Cos(CDbl(angle))) + 10# * Abs(Sin(CDbl(angle)))) / 2#
        TestAssertClose stats, "review.angle.envelope." & CStr(angle), model.ConcreteBoundaryHalfProjection(1, 1#, 0#), expected, 0.0000001
    Next angle
    TestGeometryReviewHugeAngle stats, -1E+20
    TestGeometryReviewHugeAngle stats, 1E+20
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Если тригонометрия среды не принимает огромный аргумент, модель должна отказать понятной ошибкой до добавления элемента; иначе сверяем допустимое направление с теми же Sin/Cos.
Private Sub TestGeometryReviewHugeAngle(ByRef stats As TTestStats, ByVal angle As Double)
    Dim originalSin As Double, originalCos As Double, trigFailure As Long, failure As Long
    Dim failureSource As String, reason As String, model As CSectionModel, actual As Double
    On Error Resume Next
    originalSin = Sin(angle): originalCos = Cos(angle): trigFailure = Err.Number: Err.Clear
    On Error GoTo 0
    Set model = New CSectionModel
    On Error Resume Next
    model.AddConcreteElement 0#, 0#, 200#, 1, shapeType:="Rectangle", width:=20#, height:=10#, rotation:=angle
    failure = Err.Number: failureSource = Err.Source: reason = Err.Description: Err.Clear
    On Error GoTo 0
    If trigFailure <> 0 Then
        TestAssertTrue stats, "review.angle.huge.errorCode." & CStr(angle), failure = vbObjectError + 5222
        TestAssertTrue stats, "review.angle.huge.errorSource." & CStr(angle), failureSource = "CSectionModel"
        TestAssertTrue stats, "review.angle.huge.errorReason." & CStr(angle), InStr(1, reason, "угол", vbTextCompare) > 0
        TestAssertClose stats, "review.angle.huge.noPartialElement." & CStr(angle), model.ConcreteCount, 0#, 0#
    Else
        TestAssertTrue stats, "review.angle.huge.accepted." & CStr(angle), failure = 0
        If failure <> 0 Then Exit Sub
        actual = model.ConcreteRotation(1)
        TestAssertTrue stats, "review.angle.huge.range." & CStr(angle), actual > -GEOM_PI / 2# And actual <= GEOM_PI / 2#
        TestAssertClose stats, "review.angle.huge.axisEquivalent." & CStr(angle), Cos(actual) * originalSin - Sin(actual) * originalCos, 0#, 0.00000001
        TestAssertClose stats, "review.angle.huge.envelope." & CStr(angle), model.ConcreteBoundaryHalfProjection(1, 1#, 0#), (20# * Abs(originalCos) + 10# * Abs(originalSin)) / 2#, 0.0000001
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет прежнюю арифметику только для ограниченных тестовых углов; не вызывается с заведомо зависающим аргументом.
Private Function TestGeometryReviewLegacyAngle(ByVal angle As Double) As Double
    Do While angle > GEOM_PI / 2#: angle = angle - GEOM_PI: Loop
    Do While angle <= -GEOM_PI / 2#: angle = angle + GEOM_PI: Loop
    TestGeometryReviewLegacyAngle = angle
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет устойчивость бетонных и приведённых характеристик при больших положительных и отрицательных WCS-переносах.
Private Sub TestGeometryReviewTranslationInertia(ByRef stats As TTestStats)
    Dim fixtureKind As Long, es As Variant, shift As Variant, model As CSectionModel, props As CSectionPropertiesCalculator
    Dim reference As CSectionPropertiesCalculator, area As Double, cx As Double, cy As Double, ix As Double, iy As Double, ixy As Double
    Dim prefix As String, x As Double, y As Double, cxGlobal As Double, cyGlobal As Double, coordinateTolerance As Double
    Dim referenceModel As CSectionModel, direction As Long, referenceCore(1 To 3) As Double
    Dim axisX(1 To 3) As Double, axisY(1 To 3) As Double, includeRebar As Boolean
    axisX(1) = 1#: axisY(2) = 1#: axisX(3) = 1#: axisY(3) = 1# ' Проверяем две координатные оси и диагональ, не ограничиваясь главными направлениями.
    For fixtureKind = 1 To 2
        For Each es In Array(0#, 150000#, 200000#, 220000#)
            TestGeometryReviewInertiaOracle fixtureKind, CDbl(es), area, cx, cy, ix, iy, ixy
            Set model = TestGeometryReviewInertiaModel(fixtureKind, 0#, 0#)
            Set referenceModel = model: includeRebar = CDbl(es) > 0#
            Set reference = New CSectionPropertiesCalculator
            If CDbl(es) = 0# Then reference.CalculateConcrete model Else reference.CalculateTransformedByModuli model, 30000#, CDbl(es)
            For direction = 1 To 3
                referenceCore(direction) = reference.CoreDistanceAlong(referenceModel, axisX(direction), axisY(direction), includeRebar)
            Next direction
            For Each shift In Array(0#, 1000000#, -1000000#, 1000000000#, -1000000000#, 10000000000#, -10000000000#)
                x = CDbl(shift): y = -CDbl(shift)
                Set model = TestGeometryReviewInertiaModel(fixtureKind, x, y)
                Set props = New CSectionPropertiesCalculator
                If CDbl(es) = 0# Then props.CalculateConcrete model Else props.CalculateTransformedByModuli model, 30000#, CDbl(es)
                prefix = "review.inertia.kind=" & CStr(fixtureKind) & ".Es=" & CStr(es) & ".WCS=" & CStr(shift)
                TestAssertClose stats, prefix & ".area", props.Area, area, GeomMax(0.000001, area * 0.00000000001)
                TestAssertClose stats, prefix & ".Ixc", props.Ixc, ix, GeomMax(0.000001, Abs(ix) * 0.00000000001)
                TestAssertClose stats, prefix & ".Iyc", props.Iyc, iy, GeomMax(0.000001, Abs(iy) * 0.00000000001)
                TestAssertClose stats, prefix & ".Ixyc", props.Ixyc, ixy, GeomMax(0.000001, Abs(ixy) * 0.00000000001)
                TestAssertTrue stats, prefix & ".exactCentralIx", props.Ixc = reference.Ixc
                TestAssertTrue stats, prefix & ".exactCentralIy", props.Iyc = reference.Iyc
                TestAssertTrue stats, prefix & ".exactCentralIxy", props.Ixyc = reference.Ixyc
                TestAssertTrue stats, prefix & ".exactPrincipalI1", props.PrincipalI1 = reference.PrincipalI1
                TestAssertTrue stats, prefix & ".exactPrincipalI2", props.PrincipalI2 = reference.PrincipalI2
                For direction = 1 To 3
                    TestAssertClose stats, prefix & ".coreDirection=" & CStr(direction), _
                        props.CoreDistanceAlong(model, axisX(direction), axisY(direction), includeRebar), _
                        referenceCore(direction), GeomMax(0.00000001, Abs(referenceCore(direction)) * 0.0000001) ' Проверка r допускает только субмикронный/малый относительный шум проекций в удалённой WCS; actual и разница сохраняются в отчёте.
                Next direction
                cxGlobal = x + cx: cyGlobal = y + cy
                coordinateTolerance = GeomMax(0.00000001, (Abs(x) + Abs(y)) * 0.0000000000000004) ' В WCS=1e10 мм сам Double хранит глобальную координату с шагом около 2e-6 мм; локальные I выше сравниваются точно.
                TestAssertClose stats, prefix & ".centroidX", props.CentroidX, cxGlobal, coordinateTolerance
                TestAssertClose stats, prefix & ".centroidY", props.CentroidY, cyGlobal, coordinateTolerance
                TestAssertClose stats, prefix & ".globalSx", props.StaticMomentX, area * cyGlobal, GeomMax(0.000001, Abs(area * cyGlobal) * 0.00000000001)
                TestAssertClose stats, prefix & ".globalSy", props.StaticMomentY, area * cxGlobal, GeomMax(0.000001, Abs(area * cxGlobal) * 0.00000000001)
                TestAssertClose stats, prefix & ".globalIx", props.Ix, ix + area * cyGlobal * cyGlobal, GeomMax(0.000001, Abs(ix + area * cyGlobal * cyGlobal) * 0.00000000001)
                TestAssertClose stats, prefix & ".globalIy", props.Iy, iy + area * cxGlobal * cxGlobal, GeomMax(0.000001, Abs(iy + area * cxGlobal * cxGlobal) * 0.00000000001)
                TestAssertClose stats, prefix & ".globalIxy", props.Ixy, ixy + area * cxGlobal * cyGlobal, GeomMax(0.000001, Abs(ixy + area * cxGlobal * cyGlobal) * 0.00000000001)
                If fixtureKind = 1 Then
                    TestAssertTrue stats, prefix & ".sourceX", model.ConcreteX(1) = x
                    TestAssertTrue stats, prefix & ".sourceY", model.ConcreteY(1) = y
                Else
                    TestAssertTrue stats, prefix & ".sourceX", model.ConcreteX(1) = x - 100#
                    TestAssertTrue stats, prefix & ".sourceY", model.ConcreteY(1) = y - 40#
                    TestAssertTrue stats, prefix & ".sourceRebarX", model.RebarX(1) = x - 80#
                    TestAssertTrue stats, prefix & ".sourceRebarY", model.RebarY(1) = y - 40#
                End If
            Next shift
        Next es
    Next fixtureKind
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт либо один реальный прямоугольник 300x200, либо два неравных волокна и два разных стержня без изменения их координат при расчёте.
Private Function TestGeometryReviewInertiaModel(ByVal fixtureKind As Long, ByVal shiftX As Double, ByVal shiftY As Double) As CSectionModel
    Dim model As CSectionModel
    Set model = New CSectionModel
    If fixtureKind = 1 Then
        model.AddConcreteElement shiftX, shiftY, 60000#, 1, shapeType:="Rectangle", width:=300#, height:=200#, localIx:=200000000#, localIy:=450000000#
    Else
        model.AddConcreteElement shiftX - 100#, shiftY - 40#, 20000#, 1, shapeType:="Rectangle", width:=200#, height:=100#, localIx:=20000# * 100# ^ 2 / 12#, localIy:=20000# * 200# ^ 2 / 12#
        model.AddConcreteElement shiftX + 150#, shiftY + 90#, 30000#, 1, shapeType:="Rectangle", width:=150#, height:=200#, localIx:=30000# * 200# ^ 2 / 12#, localIy:=30000# * 150# ^ 2 / 12#
        model.AddRebarElement shiftX - 80#, shiftY - 40#, 20#, 100# * GEOM_PI, "A400"
        model.AddRebarElement shiftX + 130#, shiftY + 80#, 32#, 256# * GEOM_PI, "A500"
    End If
    Set TestGeometryReviewInertiaModel = model
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимая двухпроходная аналитика в исходных локальных координатах; рабочий класс накапливает первые и вторые моменты в одном проходе.
Private Sub TestGeometryReviewInertiaOracle(ByVal fixtureKind As Long, ByVal es As Double, ByRef area As Double, _
        ByRef cx As Double, ByRef cy As Double, ByRef ix As Double, ByRef iy As Double, ByRef ixy As Double)
    Dim weight(1 To 4) As Double, x(1 To 4) As Double, y(1 To 4) As Double
    Dim ownIx(1 To 4) As Double, ownIy(1 To 4) As Double, count As Long, i As Long, factor As Double, barArea As Double
    area = 0#: cx = 0#: cy = 0#: ix = 0#: iy = 0#: ixy = 0#
    If fixtureKind = 1 Then
        count = 1: weight(1) = 60000#: ownIx(1) = 200000000#: ownIy(1) = 450000000#
    Else
        count = 2
        weight(1) = 20000#: x(1) = -100#: y(1) = -40#: ownIx(1) = 20000# * 100# ^ 2 / 12#: ownIy(1) = 20000# * 200# ^ 2 / 12#
        weight(2) = 30000#: x(2) = 150#: y(2) = 90#: ownIx(2) = 30000# * 200# ^ 2 / 12#: ownIy(2) = 30000# * 150# ^ 2 / 12#
        If es > 0# Then
            count = 4: factor = es / 30000# - 1# ' Сетка содержит бетон в месте стержня: добавочный вес равен прежнему (Es/Eb-1)*As.
            barArea = 100# * GEOM_PI: weight(3) = factor * barArea: x(3) = -80#: y(3) = -40#
            ownIx(3) = factor * barArea * barArea / (4# * GEOM_PI): ownIy(3) = ownIx(3)
            barArea = 256# * GEOM_PI: weight(4) = factor * barArea: x(4) = 130#: y(4) = 80#
            ownIx(4) = factor * barArea * barArea / (4# * GEOM_PI): ownIy(4) = ownIx(4)
        End If
    End If
    For i = 1 To count
        area = area + weight(i): cx = cx + weight(i) * x(i): cy = cy + weight(i) * y(i)
    Next i
    cx = cx / area: cy = cy / area
    For i = 1 To count
        ix = ix + ownIx(i) + weight(i) * (y(i) - cy) ^ 2
        iy = iy + ownIy(i) + weight(i) * (x(i) - cx) ^ 2
        ixy = ixy + weight(i) * (x(i) - cx) * (y(i) - cy)
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет обе явные политики касательной у стыка, близкие станции, открытые концы и прежнюю семантику замыкания.
Public Function TestRunWorkingCornerJointTests() As String
    On Error GoTo Failed
    Dim stats As TTestStats
    TestGeometryReviewCornerJoints stats
    TestAppendLine stats, "TOTAL_WORKING_CORNER_JOINTS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    TestRunWorkingCornerJointTests = stats.Report
    Exit Function
Failed:
    TestRunWorkingCornerJointTests = stats.Report & "FAIL: cornerJoint.runtime=" & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимая таблица ожидаемых сегментов различает численный шум стыка и физически разные станции вне существующего допуска.
Private Sub TestGeometryReviewCornerJoints(ByRef stats As TTestStats)
    Dim shape As Long, reversed As Variant, incoming As Variant, deltaIndex As Long, j As Long, endIndex As Long
    Dim path As CGeometryPath, edge As TRegionEdge, first As TRegionEdge, last As TRegionEdge, unchanged As Boolean
    Dim deltas As Variant, station As Double, expectedIndex As Long, normalOffset As Double, oldLength As Double, prefix As String
    deltas = Array(-2# * GEOM_TOLERANCE, -GEOM_TOLERANCE / 2#, -1E-14, 0#, 1E-14, GEOM_TOLERANCE / 2#, 2# * GEOM_TOLERANCE)
    For shape = 1 To 3
        For Each reversed In Array(False, True)
            Set path = TestGeometryReviewCornerPath(shape, CBool(reversed))
            path.CopySegment 1, first: path.CopySegment path.SegmentCount, last: oldLength = path.Length
            normalOffset = 5#
            For j = 1 To path.SegmentCount
                path.CopySegment j, edge
                If edge.Sweep <> 0# Then normalOffset = Sgn(edge.Sweep) * 5#
            Next j
            For j = 1 To path.SegmentCount - 1
                station = path.AnchorStation(path.SegmentID(j + 1), 0#)
                For Each incoming In Array(False, True)
                    For deltaIndex = 0 To 6
                        expectedIndex = j + 1
                        If CBool(incoming) And deltaIndex < 6 Then expectedIndex = j
                        If Not CBool(incoming) And deltaIndex = 0 Then expectedIndex = j
                        prefix = "review.joint.shape=" & CStr(shape) & ".reverse=" & CStr(reversed) & ".j=" & CStr(j) & ".incoming=" & CStr(incoming) & ".delta=" & CStr(deltaIndex)
                        TestGeometryReviewJointSample stats, path, station + CDbl(deltas(deltaIndex)), expectedIndex, CBool(incoming), normalOffset, prefix
                    Next deltaIndex
                Next incoming
            Next j
            For Each incoming In Array(False, True)
                For endIndex = 0 To 1
                    station = 0#: expectedIndex = 1
                    If endIndex = 1 Then station = path.Length: expectedIndex = path.SegmentCount
                    prefix = "review.jointEndpoint.shape=" & CStr(shape) & ".reverse=" & CStr(reversed) & ".incoming=" & CStr(incoming) & ".end=" & CStr(endIndex)
                    TestGeometryReviewJointSample stats, path, station, expectedIndex, CBool(incoming), normalOffset, prefix
                Next endIndex
            Next incoming
            prefix = "review.jointSource.shape=" & CStr(shape) & ".reverse=" & CStr(reversed)
            TestAssertClose stats, prefix & ".length", path.Length, oldLength, 0#
            path.CopySegment 1, edge
            unchanged = edge.X1 = first.X1 And edge.Y1 = first.Y1 And edge.X2 = first.X2 And edge.Y2 = first.Y2
            path.CopySegment path.SegmentCount, edge
            TestAssertTrue stats, prefix & ".endpoints", unchanged And edge.X1 = last.X1 And edge.Y1 = last.Y1 And edge.X2 = last.X2 And edge.Y2 = last.Y2
        Next reversed
    Next shape
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Один острый стык прямых, гладкий стык прямой с дугой и замкнутый квадрат дают различимые касательные и проверку выбора дуги для нормального переноса.
Private Function TestGeometryReviewCornerPath(ByVal shape As Long, ByVal reversed As Boolean) As CGeometryPath
    Dim path As CGeometryPath, i As Long, edges() As TRegionEdge
    Set path = New CGeometryPath: path.Initialize "JointPath"
    If shape = 1 Then
        path.AddLine "First", 0#, 0#, 10#, 0#
        path.AddLine "Second", 10#, 0#, 10#, 20#
        path.Seal
    ElseIf shape = 2 Then
        path.AddLine "Line", 0#, 0#, 10#, 0#
        path.AddArc "Arc", 10#, 5#, 5#, -GEOM_PI / 2#, GEOM_PI / 2#
        path.Seal
    Else
        TestGeometryReviewSquareEdges edges, 10#, 0#, 0#, False
        For i = 1 To 4: path.AddLine "Edge" & CStr(i), edges(i).X1, edges(i).Y1, edges(i).X2, edges(i).Y2: Next i
        path.Seal True
    End If
    If reversed Then Set path = path.Reversed
    Set TestGeometryReviewCornerPath = path
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ожидаемый сегмент задаётся тестовой таблицей, без вызова рабочего поиска; AtDistance и ValidateNormalTransfer обязаны выбрать один и тот же сегмент.
Private Sub TestGeometryReviewJointSample(ByRef stats As TTestStats, ByVal path As CGeometryPath, ByVal station As Double, _
        ByVal expectedIndex As Long, ByVal incoming As Boolean, ByVal normalOffset As Double, ByVal prefix As String)
    Dim edge As TRegionEdge, first As Double, last As Double, parameter As Double, expectedX As Double, expectedY As Double
    Dim expectedTx As Double, expectedTy As Double, norm As Double, x As Double, y As Double, tx As Double, ty As Double, failure As Long
    path.CopySegment expectedIndex, edge
    first = path.AnchorStation(path.SegmentID(expectedIndex), 0#): last = path.AnchorStation(path.SegmentID(expectedIndex), 1#)
    parameter = GeomMax(0#, GeomMin(1#, (station - first) / (last - first)))
    modGeometryPrimitives.EdgePoint edge, parameter, expectedX, expectedY
    modGeometryPrimitives.EdgeTangent edge, parameter, expectedTx, expectedTy
    norm = Sqr(expectedTx * expectedTx + expectedTy * expectedTy): expectedTx = expectedTx / norm: expectedTy = expectedTy / norm
    path.AtDistance station, x, y, tx, ty, incoming
    TestAssertClose stats, prefix & ".x", x, expectedX, 0.0000000001
    TestAssertClose stats, prefix & ".y", y, expectedY, 0.0000000001
    TestAssertClose stats, prefix & ".tangentX", tx, expectedTx, 0.000000000001
    TestAssertClose stats, prefix & ".tangentY", ty, expectedTy, 0.000000000001
    On Error Resume Next
    path.ValidateNormalTransfer station, normalOffset, incoming: failure = Err.Number: Err.Clear
    On Error GoTo 0
    TestAssertTrue stats, prefix & ".normalTransferSameSegment", (failure <> 0) = (edge.Sweep <> 0#)
End Sub
