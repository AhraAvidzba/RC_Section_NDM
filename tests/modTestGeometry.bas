Attribute VB_Name = "modTestGeometry"
Option Explicit

' ==========================================================================
' Тесты геометрии, сетки и раскладки арматуры
' ==========================================================================
' Модуль защищает договоренности по Circle, RoundedRectangle и LShape: габариты,
' дискретизацию, автоматическую арматуру и semantic-аннотации для схемы.

Private Type TTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' Запускает связанный набор операций и возвращает пользователю итоговый статус выполнения.
Public Function RunGeometryTests() As String
    On Error GoTo Failed

    Dim stats As TTestStats

    TestRectangle stats
    TestSymmetricRoundedRectangle stats
    TestAsymmetricRadii stats
    TestCircleGeometry stats
    TestCircleInvalidData stats
    TestCircleAutoRebarLayout stats
    TestLShapeGeometry stats
    TestLShapeAutoRebarLayout stats
    TestLShapeSeparateLineOffsets stats
    TestLShapeAdditionalRebarRows stats
    TestSectionModelFromGeneratedGeometry stats
    TestRebarAnnotationAnchors stats
    TestAutoCADImporterBuildsSectionModel stats
    TestInvalidData stats
    TestBoundarySubcellMesh stats
    TestMeshConvergence stats
    TestPerformance stats

    AppendLine stats, "TOTAL: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunGeometryTests = stats.Report
    Exit Function

Failed:
    RunGeometryTests = "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestSectionModelFromGeneratedGeometry(ByRef stats As TTestStats)
    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter 300#

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 50#, 50#, 1, 2

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "CircleSourceName-1", 100#, 0#, 20#, 0#, "A400", "source name", geom
    rebars.AddBar "CircleSourceName-2", -100#, 0#, 20#, 0#, "A400", "source name", geom

    Dim model As CSectionModel
    Set model = BuildGeneratedSectionModel(mesh, rebars, "Circle")

    AssertTrue stats, "model.concrete.count", model.ConcreteCount = mesh.FiberCount
    AssertTrue stats, "model.rebar.count", model.RebarCount = rebars.Count
    AssertTrue stats, "model.concrete.id", model.ConcreteID(1) = "C1"
    AssertTrue stats, "model.rebar.id", model.RebarID(1) = "R1"
    AssertTrue stats, "model.rebar.sourceName", model.RebarSourceName(1) = "CircleSourceName-1"
    AssertClose stats, "model.rebar.area", model.RebarArea(1), GEOM_PI * 20# * 20# / 4#, 0.000000001
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestRebarAnnotationAnchors(ByRef stats As TTestStats)
    Dim circleBuilder As CCircleRebarLayoutBuilder
    Set circleBuilder = New CCircleRebarLayoutBuilder

    Dim circleBars As CRebarLayout
    Set circleBars = circleBuilder.Build(300#, 10#, -20#, 40#, 8, 20#, "A400")
    AssertTrue stats, "annotation.circle.count", circleBars.AnnotationCount = 1
    AssertTrue stats, "annotation.circle.group", circleBars.AnnotationGroupName(1) = "Circle"
    AssertClose stats, "annotation.circle.startX", circleBars.AnnotationStartX(1), -100#, 0.000001
    AssertClose stats, "annotation.circle.endX", circleBars.AnnotationEndX(1), 120#, 0.000001
    AssertClose stats, "annotation.circle.normalY", circleBars.AnnotationNormalY(1), 1#, 0.000001
    AssertClose stats, "annotation.circle.axisDistance", circleBars.AnnotationAxisDistance(1), 40#, 0.000001

    Dim lshapeBuilder As CLShapeRebarLayoutBuilder
    Set lshapeBuilder = New CLShapeRebarLayoutBuilder
    Dim lshapeBars As CRebarLayout
    Set lshapeBars = lshapeBuilder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        Array(40#, 40#, 20#, 20#, 3, 0, 50#, 70#, 0#, 0#, 0#, 0#, 0#, 0#, "Stacked", "Stacked"), _
        EmptyFaceSettings(), EmptyFaceSettings(), EmptyFaceSettings(), "A400")

    AssertTrue stats, "annotation.lshape.count", lshapeBars.AnnotationCount = 1
    AssertTrue stats, "annotation.lshape.group", lshapeBars.AnnotationGroupName(1) = "H1.as_1"
    AssertClose stats, "annotation.lshape.startY", lshapeBars.AnnotationStartY(1), 300#, 0.000001
    AssertClose stats, "annotation.lshape.endY", lshapeBars.AnnotationEndY(1), 730#, 0.000001
    AssertClose stats, "annotation.lshape.normalX", lshapeBars.AnnotationNormalX(1), -1#, 0.000001
    AssertClose stats, "annotation.lshape.axisDistance", lshapeBars.AnnotationAxisDistance(1), 40#, 0.000001

    Dim noBars As CRebarLayout
    Set noBars = New CRebarLayout
    noBars.AddAnnotationAnchor "SyntheticOnly", 0#, 0#, 1#, 0#, 0#, 1#
    AssertTrue stats, "annotation.manual.anchor.allowed", noBars.AnnotationCount = 1

    Dim zeroFaceBuilder As CLShapeRebarLayoutBuilder
    Set zeroFaceBuilder = New CLShapeRebarLayoutBuilder
    Dim zeroFaceBars As CRebarLayout
    Set zeroFaceBars = zeroFaceBuilder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        EmptyFaceSettings(), _
        EmptyFaceSettings(), _
        Array(40#, 40#, 20#, 20#, 0, 0, 50#, 70#, 0#, 0#, 0#, 0#, 0#, 0#, "Stacked", "Stacked"), _
        Array(40#, 40#, 20#, 20#, 2, 0, 50#, 70#, 0#, 0#, 0#, 0#, 0#, 0#, "Stacked", "Stacked"), _
        "A400")
    AssertTrue stats, "annotation.zeroFace.hasOtherBars", zeroFaceBars.Count = 2
    AssertTrue stats, "annotation.zeroFace.noAnchor", Not HasAnnotationGroup(zeroFaceBars, "B1.as_1")
    AssertTrue stats, "annotation.zeroFace.otherAnchor", HasAnnotationGroup(zeroFaceBars, "B2.as_1")

    Dim zeroCircleBars As CRebarLayout
    Set zeroCircleBars = circleBuilder.Build(300#, 10#, -20#, 40#, 0, 20#, "A400", 20#, 30#, "Stacked", "Stacked")
    AssertTrue stats, "annotation.circle.zeroCount.noBars", zeroCircleBars.Count = 0
    AssertTrue stats, "annotation.circle.zeroCount.noAnchor", zeroCircleBars.AnnotationCount = 0

    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter 300#, 10#, -20#
    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 50#, 50#, 1, 2

    Dim model As CSectionModel
    Set model = BuildGeneratedSectionModel(mesh, circleBars, "Circle")
    Dim annotationBuilder As CCircleAnnotationBuilder
    Set annotationBuilder = New CCircleAnnotationBuilder
    annotationBuilder.Build model, geom, circleBars
    AssertTrue stats, "annotation.model.count", model.AnnotationCount = 3
    AssertTrue stats, "annotation.model.rebarLabel", HasSectionAnnotation(model, "REBAR_ANNOTATION", "REBAR_Circle")
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestAutoCADImporterBuildsSectionModel(ByRef stats As TTestStats)
    Dim concreteRegions(1 To 2, 1 To 7) As Variant
    concreteRegions(1, 1) = 0.0000000001
    concreteRegions(1, 2) = 0#
    concreteRegions(1, 3) = 0#
    concreteRegions(2, 1) = 200#
    concreteRegions(2, 2) = 10#
    concreteRegions(2, 3) = 20#
    concreteRegions(2, 4) = 600#
    concreteRegions(2, 5) = 150#
    concreteRegions(2, 6) = 0#
    concreteRegions(2, 7) = "ABC"

    Dim rebarRegions(1 To 1, 1 To 4) As Variant
    rebarRegions(1, 1) = GEOM_PI * 20# * 20# / 4#
    rebarRegions(1, 2) = -30#
    rebarRegions(1, 3) = 40#
    rebarRegions(1, 4) = "DEF"

    Dim importer As CAutoCADSectionModelImporter
    Set importer = New CAutoCADSectionModelImporter

    Dim model As CSectionModel
    Set model = importer.BuildFromRegionArrays(concreteRegions, rebarRegions, "Rebar", 0.000001)

    AssertTrue stats, "autocad.import.source", model.SourceType = "AutoCADImport"
    AssertTrue stats, "autocad.import.concrete.count", model.ConcreteCount = 1
    AssertTrue stats, "autocad.import.rebar.count", model.RebarCount = 1
    AssertTrue stats, "autocad.import.concrete.id", model.ConcreteID(1) = "C1"
    AssertClose stats, "autocad.import.concrete.x", model.ConcreteX(1), 10#, 0.000001
    AssertClose stats, "autocad.import.concrete.y", model.ConcreteY(1), 20#, 0.000001
    AssertClose stats, "autocad.import.concrete.width", model.ConcreteWidth(1), 3#, 0.000001
    AssertClose stats, "autocad.import.concrete.height", model.ConcreteHeight(1), 6#, 0.000001
    AssertClose stats, "autocad.import.concrete.localIx", model.ConcreteLocalIx(1), 600#, 0.000001
    AssertClose stats, "autocad.import.concrete.localIy", model.ConcreteLocalIy(1), 150#, 0.000001
    AssertClose stats, "autocad.import.rebar.diameter", model.RebarDiameter(1), 20#, 0.000001
    AssertTrue stats, "autocad.import.rebar.marker", model.RebarSteelClass(1) = "Rebar"
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestRectangle(ByRef stats As TTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 200#, 100#, 0#, 0#, 0#, 0#

    Dim props As CSectionPropertiesCalculator
    Set props = MeshProps(geom, 2.5, 2.5)

    AssertClose stats, "rect.area", props.Area, 200# * 100#, 0.000001
    AssertClose stats, "rect.cx", props.CentroidX, 0#, 0.000001
    AssertClose stats, "rect.cy", props.CentroidY, 0#, 0.000001
    AssertRelative stats, "rect.Ix", props.Ixc, 200# * 100# ^ 3 / 12#, 0.001
    AssertRelative stats, "rect.Iy", props.Iyc, 100# * 200# ^ 3 / 12#, 0.001
    AssertClose stats, "rect.Ixy", props.Ixyc, 0#, 0.000001
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestLShapeAdditionalRebarRows(ByRef stats As TTestStats)
    Dim builder As CLShapeRebarLayoutBuilder
    Set builder = New CLShapeRebarLayoutBuilder

    Dim stacked As CRebarLayout
    Set stacked = builder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        Array(50#, 50#, 20#, 20#, 1, 0, 100#, 100#, 100#, 100#, 20#, 0#, 30#, 0#, "Stacked", "Stacked"), _
        EmptyFaceSettings(), EmptyFaceSettings(), EmptyFaceSettings(), "A400")

    AssertTrue stats, "lshape.rows.stacked.count", stacked.Count = 3
    AssertClose stats, "lshape.rows.stacked.row1.x", stacked.X(1), 50#, 0.000001
    AssertClose stats, "lshape.rows.stacked.row2.x", stacked.X(2), 70#, 0.000001
    AssertClose stats, "lshape.rows.stacked.row3.x", stacked.X(3), 95#, 0.000001
    AssertClose stats, "lshape.rows.stacked.sameY", stacked.Y(3), stacked.Y(1), 0.000001

    Dim sideVertical As CRebarLayout
    Set sideVertical = builder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        Array(50#, 50#, 20#, 20#, 1, 0, 100#, 100#, 100#, 100#, 20#, 0#, 30#, 0#, "SideBySide", "SideBySide"), _
        EmptyFaceSettings(), EmptyFaceSettings(), EmptyFaceSettings(), "A400")

    AssertTrue stats, "lshape.rows.side.vertical.count", sideVertical.Count = 3
    AssertClose stats, "lshape.rows.side.vertical.sameX", sideVertical.X(3), sideVertical.X(1), 0.000001
    AssertClose stats, "lshape.rows.side.vertical.row2.y", sideVertical.Y(2), sideVertical.Y(1) - 20#, 0.000001
    AssertClose stats, "lshape.rows.side.vertical.row3.y", sideVertical.Y(3), sideVertical.Y(1) - 45#, 0.000001

    Dim sideHorizontal As CRebarLayout
    Set sideHorizontal = builder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        EmptyFaceSettings(), EmptyFaceSettings(), _
        Array(50#, 50#, 20#, 20#, 1, 0, 50#, 50#, 50#, 50#, 20#, 0#, 30#, 0#, "SideBySide", "SideBySide"), _
        EmptyFaceSettings(), "A400")

    AssertTrue stats, "lshape.rows.side.horizontal.count", sideHorizontal.Count = 3
    AssertClose stats, "lshape.rows.side.horizontal.row2.x", sideHorizontal.X(2), sideHorizontal.X(1) + 20#, 0.000001
    AssertClose stats, "lshape.rows.side.horizontal.row3.x", sideHorizontal.X(3), sideHorizontal.X(1) + 45#, 0.000001
    AssertClose stats, "lshape.rows.side.horizontal.sameY", sideHorizontal.Y(3), sideHorizontal.Y(1), 0.000001

    Dim thirdOnly As CRebarLayout
    Set thirdOnly = builder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        Array(50#, 50#, 20#, 20#, 1, 0, 100#, 100#, 100#, 100#, 0#, 0#, 30#, 0#, "Stacked", "Stacked"), _
        EmptyFaceSettings(), EmptyFaceSettings(), EmptyFaceSettings(), "A400")

    AssertTrue stats, "lshape.rows.thirdOnly.count", thirdOnly.Count = 2
    AssertClose stats, "lshape.rows.thirdOnly.row3.x", thirdOnly.X(2), thirdOnly.X(1) + 25#, 0.000001
    AssertClose stats, "lshape.rows.thirdOnly.sameY", thirdOnly.Y(2), thirdOnly.Y(1), 0.000001

    Dim noFirstRow As CRebarLayout
    Set noFirstRow = builder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        Array(50#, 50#, 20#, 20#, 0, 0, 100#, 100#, 100#, 100#, 20#, 20#, 30#, 30#, "Stacked", "Stacked"), _
        Array(50#, 50#, 20#, 20#, 1, 0, 100#, 100#, 100#, 100#, 0#, 0#, 0#, 0#, "Stacked", "Stacked"), _
        EmptyFaceSettings(), EmptyFaceSettings(), "A400")

    AssertTrue stats, "lshape.rows.noFirstRow.ignored", noFirstRow.Count = 1
    AssertTrue stats, "lshape.rows.noFirstRow.source", InStr(1, noFirstRow.BarID(1), "H2", vbTextCompare) > 0

    Dim mixed As CRebarLayout
    Set mixed = builder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        Array(50#, 50#, 20#, 20#, 1, 0, 100#, 100#, 100#, 100#, 20#, 0#, 30#, 0#, "SideBySide", "Stacked"), _
        EmptyFaceSettings(), EmptyFaceSettings(), EmptyFaceSettings(), "A400")

    AssertTrue stats, "lshape.rows.mixed.count", mixed.Count = 3
    AssertClose stats, "lshape.rows.mixed.row2.y", mixed.Y(2), mixed.Y(1) - 20#, 0.000001
    AssertClose stats, "lshape.rows.mixed.row3.x", mixed.X(3), mixed.X(1) + 25#, 0.000001
    AssertClose stats, "lshape.rows.mixed.row3.y", mixed.Y(3), mixed.Y(1), 0.000001

    Dim separateLineDiameters As CRebarLayout
    Set separateLineDiameters = builder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        Array(50#, 50#, 20#, 20#, 1, 1, 100#, 100#, 100#, 100#, 18#, 22#, 28#, 34#, "Stacked", "Stacked"), _
        EmptyFaceSettings(), EmptyFaceSettings(), EmptyFaceSettings(), "A400")

    AssertTrue stats, "lshape.rows.separateLineDiameters.count", separateLineDiameters.Count = 6
    AssertClose stats, "lshape.rows.separateLineDiameters.row2_1.d", separateLineDiameters.Diameter(2), 18#, 0.000001
    AssertClose stats, "lshape.rows.separateLineDiameters.row3_1.d", separateLineDiameters.Diameter(3), 28#, 0.000001
    AssertClose stats, "lshape.rows.separateLineDiameters.row2_2.d", separateLineDiameters.Diameter(5), 22#, 0.000001
    AssertClose stats, "lshape.rows.separateLineDiameters.row3_2.d", separateLineDiameters.Diameter(6), 34#, 0.000001

    Dim everySecond As CRebarLayout
    Set everySecond = builder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        Array(50#, 50#, 20#, 20#, 5, 0, 50#, 50#, 50#, 50#, 20#, 0#, 30#, 0#, "Stacked", "Stacked", "EverySecondBar", "EachBar"), _
        EmptyFaceSettings(), EmptyFaceSettings(), EmptyFaceSettings(), "A400")

    AssertTrue stats, "lshape.rows.everySecond.count", everySecond.Count = 13
    AssertTrue stats, "lshape.rows.everySecond.row2.partial", CountBarsWithRow(everySecond, "row_2") = 3
    AssertTrue stats, "lshape.rows.everySecond.row3.each", CountBarsWithRow(everySecond, "row_3") = 5

    AssertLShapeRebarRowsError stats, "lshape.rows.invalid.location", _
        Array(50#, 50#, 20#, 20#, 1, 0, 100#, 100#, 100#, 100#, 20#, 0#, 0#, 0#, "Diagonal", "Stacked")
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestLShapeGeometry(ByRef stats As TTestStats)
    Dim geom As CGeometryLShape
    Set geom = New CGeometryLShape
    geom.Initialize 250#, 550#, 600#, 250#, 10#, -20#

    Dim available As Boolean
    Dim analyticalArea As Double
    Dim centroidX As Double
    Dim centroidY As Double
    analyticalArea = geom.AnalyticalArea(available)
    geom.AnalyticalCentroid available, centroidX, centroidY

    AssertTrue stats, "lshape.area.available", available
    AssertClose stats, "lshape.area.analytical", analyticalArea, 600# * 250# + 250# * 550#, 0.000001
    AssertClose stats, "lshape.cx.analytical", centroidX, 226.304347826087, 0.000001
    AssertClose stats, "lshape.cy.analytical", centroidY, 296.304347826087, 0.000001
    AssertTrue stats, "lshape.contains.lower", geom.ContainsPoint(580#, 20#)
    AssertTrue stats, "lshape.contains.vertical", geom.ContainsPoint(100#, 700#)
    AssertTrue stats, "lshape.excludes.cutout", Not geom.ContainsPoint(500#, 700#)

    Dim centerMesh As CFiberMeshBuilder
    Set centerMesh = New CFiberMeshBuilder
    centerMesh.BuildMesh geom, 80#, 80#, 1

    Dim subcellMesh As CFiberMeshBuilder
    Set subcellMesh = New CFiberMeshBuilder
    subcellMesh.BuildMesh geom, 80#, 80#, 1, 4

    Dim centerProps As CSectionPropertiesCalculator
    Set centerProps = New CSectionPropertiesCalculator
    centerProps.CalculateConcrete BuildGeneratedSectionModel(centerMesh, Nothing)

    Dim subcellProps As CSectionPropertiesCalculator
    Set subcellProps = New CSectionPropertiesCalculator
    subcellProps.CalculateConcrete BuildGeneratedSectionModel(subcellMesh, Nothing)

    AssertTrue stats, "lshape.boundary.subcell.fibers.more", subcellMesh.FiberCount > centerMesh.FiberCount
    AssertTrue stats, "lshape.boundary.subcell.has.small.fibers", MeshHasSmallFibers(subcellMesh, 80#)
    AssertRelative stats, "lshape.boundary.subcell.area", subcellProps.Area, analyticalArea, 0.04
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestLShapeAutoRebarLayout(ByRef stats As TTestStats)
    Dim builder As CLShapeRebarLayoutBuilder
    Set builder = New CLShapeRebarLayoutBuilder

    Dim layout As CRebarLayout
    Set layout = builder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        Array(50#, 50#, 20#, 20#, 3, 0, 100#, 100#, 100#, 100#), _
        Array(50#, 50#, 20#, 20#, 0, 0, 50#, 50#, 50#, 50#), _
        Array(50#, 50#, 20#, 20#, 2, 0, 60#, 60#, 60#, 60#), _
        Array(50#, 50#, 20#, 20#, 3, 0, 100#, 100#, 100#, 100#), _
        "A400")

    AssertTrue stats, "lshape.rebar.count", layout.Count = 8
    AssertTrue stats, "lshape.rebar.requested", builder.RequestedCount = 8
    AssertClose stats, "lshape.rebar.perimeter", builder.Perimeter, 3300#, 0.000001
    AssertClose stats, "lshape.rebar.step", builder.StepAlong, 200#, 0.000001
    AssertClose stats, "lshape.rebar.firstX", layout.X(1), 50#, 0.000001
    AssertClose stats, "lshape.rebar.firstY", layout.Y(1), 525#, 0.000001
    AssertTrue stats, "lshape.rebar.zeroFaceSkipped", InStr(1, layout.BarID(1), "H2", vbTextCompare) = 0

    Dim noLineByDiameter As CRebarLayout
    Set noLineByDiameter = builder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        Array(50#, 50#, 0#, 20#, 3, 0, 100#, 100#, 100#, 100#), _
        Array(50#, 50#, 20#, 20#, 0, 0, 50#, 50#, 50#, 50#), _
        Array(50#, 50#, 20#, 20#, 2, 0, 60#, 60#, 60#, 60#), _
        Array(50#, 50#, 20#, 20#, 3, 0, 100#, 100#, 100#, 100#), _
        "A400")
    AssertTrue stats, "lshape.rebar.zeroDiameterSkipped", noLineByDiameter.Count = 5

    Dim noBarsByDiameter As CRebarLayout
    Set noBarsByDiameter = builder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        Array(50#, 50#, 0#, 0#, 3, 2, 100#, 100#, 100#, 100#), _
        Array(50#, 50#, 0#, 0#, 2, 2, 50#, 50#, 50#, 50#), _
        Array(50#, 50#, 0#, 0#, 2, 2, 60#, 60#, 60#, 60#), _
        Array(50#, 50#, 0#, 0#, 3, 3, 100#, 100#, 100#, 100#), _
        "A400")
    AssertTrue stats, "lshape.rebar.allZeroDiameters.noError", noBarsByDiameter.Count = 0

    Dim geom As CGeometryLShape
    Set geom = New CGeometryLShape
    geom.Initialize 250#, 550#, 600#, 250#

    Dim i As Long
    For i = 1 To layout.Count
        AssertTrue stats, "lshape.rebar.inside." & CStr(i), geom.ContainsPoint(layout.X(i), layout.Y(i))
    Next i

    AssertLShapeRebarError stats, "lshape.rebar.invalid.n", 600#, 550#, 250#, 250#, -1, 50#, 20#
    AssertLShapeRebarError stats, "lshape.rebar.invalid.as.small", 600#, 550#, 250#, 250#, 2, 10#, 20#
    AssertLShapeRebarError stats, "lshape.rebar.invalid.offsets", 600#, 550#, 250#, 250#, 2, 50#, 20#, 500#, 500#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestLShapeSeparateLineOffsets(ByRef stats As TTestStats)
    Dim builder As CLShapeRebarLayoutBuilder
    Set builder = New CLShapeRebarLayoutBuilder

    Dim layout As CRebarLayout
    Set layout = builder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        Array(40#, 60#, 20#, 20#, 2, 2, 80#, 120#, 30#, 70#), _
        Array(40#, 40#, 20#, 20#, 0, 0, 40#, 40#, 40#, 40#), _
        Array(40#, 40#, 20#, 20#, 0, 0, 40#, 40#, 40#, 40#), _
        Array(40#, 40#, 20#, 20#, 0, 0, 40#, 40#, 40#, 40#), _
        "A400")

    AssertTrue stats, "lshape.offsets.separate.count", layout.Count = 4
    AssertClose stats, "lshape.offsets.as_1.firstY", layout.Y(1), 680#, 0.000001
    AssertClose stats, "lshape.offsets.as_1.edgeY", layout.Y(2), 330#, 0.000001
    AssertClose stats, "lshape.offsets.as_2.firstY", layout.Y(3), 320#, 0.000001
    AssertClose stats, "lshape.offsets.as_2.edgeY", layout.Y(4), 770#, 0.000001
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestCircleAutoRebarLayout(ByRef stats As TTestStats)
    Dim builder As CCircleRebarLayoutBuilder
    Set builder = New CCircleRebarLayoutBuilder

    Dim layout As CRebarLayout
    Set layout = builder.Build(300#, 10#, -20#, 40#, 8, 20#, "A400")

    AssertTrue stats, "circle.rebar.count", layout.Count = 8
    AssertClose stats, "circle.rebar.angleStep", builder.AngleStep, 2# * GEOM_PI / 8#, 0.000000001
    AssertClose stats, "circle.rebar.axisRadius", builder.AxisRadius, 110#, 0.000001
    AssertClose stats, "circle.rebar.firstX", layout.X(1), 120#, 0.000001
    AssertClose stats, "circle.rebar.firstY", layout.Y(1), -20#, 0.000001
    AssertClose stats, "circle.rebar.centroidX", builder.CentroidX, 10#, 0.000001
    AssertClose stats, "circle.rebar.centroidY", builder.CentroidY, -20#, 0.000001
    AssertRelative stats, "circle.rebar.area", layout.Area(1), GEOM_PI * 20# * 20# / 4#, 0.000000001

    Dim stacked As CRebarLayout
    Set stacked = builder.Build(300#, 10#, -20#, 40#, 8, 20#, "A400", 20#, 30#, "Stacked", "Stacked")
    AssertTrue stats, "circle.rebar.rows.stacked.count", stacked.Count = 24
    AssertClose stats, "circle.rebar.rows.stacked.row2.x", stacked.X(2), 100#, 0.000001
    AssertClose stats, "circle.rebar.rows.stacked.row2.y", stacked.Y(2), -20#, 0.000001
    AssertClose stats, "circle.rebar.rows.stacked.row3.x", stacked.X(3), 75#, 0.000001
    AssertClose stats, "circle.rebar.rows.stacked.row3.y", stacked.Y(3), -20#, 0.000001

    Dim sideBySide As CRebarLayout
    Set sideBySide = builder.Build(300#, 10#, -20#, 40#, 8, 20#, "A400", 20#, 30#, "SideBySide", "SideBySide")
    AssertTrue stats, "circle.rebar.rows.side.count", sideBySide.Count = 24
    AssertClose stats, "circle.rebar.rows.side.row2.x", sideBySide.X(2), 120#, 0.000001
    AssertClose stats, "circle.rebar.rows.side.row2.y", sideBySide.Y(2), -40#, 0.000001
    AssertClose stats, "circle.rebar.rows.side.row3.x", sideBySide.X(3), 120#, 0.000001
    AssertClose stats, "circle.rebar.rows.side.row3.y", sideBySide.Y(3), -65#, 0.000001

    Dim mixed As CRebarLayout
    Set mixed = builder.Build(300#, 10#, -20#, 40#, 8, 20#, "A400", 20#, 30#, "SideBySide", "Stacked")
    AssertTrue stats, "circle.rebar.rows.mixed.count", mixed.Count = 24
    AssertClose stats, "circle.rebar.rows.mixed.row2.y", mixed.Y(2), -40#, 0.000001
    AssertClose stats, "circle.rebar.rows.mixed.row3.x", mixed.X(3), 95#, 0.000001
    AssertClose stats, "circle.rebar.rows.mixed.row3.y", mixed.Y(3), -20#, 0.000001

    Dim thirdOnly As CRebarLayout
    Set thirdOnly = builder.Build(300#, 10#, -20#, 40#, 8, 20#, "A400", 0#, 30#, "Stacked", "Stacked")
    AssertTrue stats, "circle.rebar.rows.thirdOnly.count", thirdOnly.Count = 16
    AssertClose stats, "circle.rebar.rows.thirdOnly.row3.x", thirdOnly.X(2), 95#, 0.000001
    AssertClose stats, "circle.rebar.rows.thirdOnly.row3.y", thirdOnly.Y(2), -20#, 0.000001

    Dim noFirstByCount As CRebarLayout
    Set noFirstByCount = builder.Build(300#, 10#, -20#, 40#, 0, 20#, "A400", 20#, 30#, "Stacked", "Stacked")
    AssertTrue stats, "circle.rebar.rows.noFirstByCount.ignored", noFirstByCount.Count = 0
    AssertClose stats, "circle.rebar.rows.noFirstByCount.angle", builder.AngleStep, 0#, 0.000000001

    Dim noFirstByDiameter As CRebarLayout
    Set noFirstByDiameter = builder.Build(300#, 10#, -20#, 40#, 8, 0#, "A400", 20#, 30#, "SideBySide", "SideBySide")
    AssertTrue stats, "circle.rebar.rows.noFirstByDiameter.ignored", noFirstByDiameter.Count = 0

    AssertCircleRebarError stats, "circle.rebar.invalid.D", 0#, 40#, 8, 20#
    AssertCircleRebarError stats, "circle.rebar.invalid.n", 300#, 40#, 1, 20#
    AssertCircleRebarError stats, "circle.rebar.invalid.ds", 300#, 40#, 8, -1#
    AssertCircleRebarError stats, "circle.rebar.invalid.as.small", 300#, 10#, 8, 20#
    AssertCircleRebarError stats, "circle.rebar.invalid.as.large", 300#, 150#, 8, 20#
    AssertCircleRebarError stats, "circle.rebar.invalid.loc2", 300#, 40#, 8, 20#, 20#, 0#, "Diagonal", "Stacked"
    AssertCircleRebarError stats, "circle.rebar.invalid.row2.outside", 300#, 20#, 8, 20#, 40#, 0#, "SideBySide", "Stacked"
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestBoundarySubcellMesh(ByRef stats As TTestStats)
    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByRadius 95#, 0#, 0#

    Dim centerMesh As CFiberMeshBuilder
    Set centerMesh = New CFiberMeshBuilder
    centerMesh.BuildMesh geom, 40#, 40#, 1

    Dim subcellMesh As CFiberMeshBuilder
    Set subcellMesh = New CFiberMeshBuilder
    subcellMesh.BuildMesh geom, 40#, 40#, 1, 4

    Dim centerProps As CSectionPropertiesCalculator
    Set centerProps = New CSectionPropertiesCalculator
    centerProps.CalculateConcrete BuildGeneratedSectionModel(centerMesh, Nothing)

    Dim subcellProps As CSectionPropertiesCalculator
    Set subcellProps = New CSectionPropertiesCalculator
    subcellProps.CalculateConcrete BuildGeneratedSectionModel(subcellMesh, Nothing)

    Dim targetArea As Double
    targetArea = GEOM_PI * 95# * 95#

    AssertTrue stats, "boundary.subcell.fibers.more", subcellMesh.FiberCount > centerMesh.FiberCount
    AssertTrue stats, "boundary.subcell.area.better", Abs(subcellProps.Area - targetArea) < Abs(centerProps.Area - targetArea)
    AssertTrue stats, "boundary.subcell.has.small.fibers", MeshHasSmallFibers(subcellMesh, 40#)
    AssertClose stats, "boundary.subcell.small.width", FirstSmallFiberWidth(subcellMesh, 40#), 10#, 0.000001
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestCircleGeometry(ByRef stats As TTestStats)
    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter 300#, 10#, -20#

    Dim available As Boolean
    Dim analyticalArea As Double
    Dim centroidX As Double
    Dim centroidY As Double
    analyticalArea = geom.AnalyticalArea(available)
    geom.AnalyticalCentroid available, centroidX, centroidY

    Dim props As CSectionPropertiesCalculator
    Set props = MeshProps(geom, 2.5, 2.5)

    Dim analyticalI As Double
    analyticalI = GEOM_PI * geom.Radius ^ 4 / 4#

    AssertTrue stats, "circle.area.available", available
    AssertRelative stats, "circle.area", props.Area, analyticalArea, 0.003
    AssertClose stats, "circle.cx", props.CentroidX, centroidX, 0.05
    AssertClose stats, "circle.cy", props.CentroidY, centroidY, 0.05
    AssertRelative stats, "circle.Ix", props.Ixc, analyticalI, 0.006
    AssertRelative stats, "circle.Iy", props.Iyc, analyticalI, 0.006
    AssertClose stats, "circle.Ixy", props.Ixyc, 0#, 0.000001 * props.Area * geom.Radius
    AssertClose stats, "circle.principal.angle", props.PrincipalAngleRad, 0#, 0.000001
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestCircleInvalidData(ByRef stats As TTestStats)
    Dim geom As CGeometryCircle
    Dim message As String
    Set geom = New CGeometryCircle
    geom.InitializeByRadius 0#
    AssertTrue stats, "circle.invalid.zeroRadius", Not geom.IsValid(message)

    Set geom = New CGeometryCircle
    geom.InitializeByDiameter -100#
    AssertTrue stats, "circle.invalid.negativeDiameter", Not geom.IsValid(message)
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestSymmetricRoundedRectangle(ByRef stats As TTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 300#, 160#, 25#, 25#, 25#, 25#

    Dim areaAvailable As Boolean
    Dim analyticalArea As Double
    analyticalArea = geom.AnalyticalArea(areaAvailable)

    Dim props As CSectionPropertiesCalculator
    Set props = MeshProps(geom, 2.5, 2.5)

    AssertTrue stats, "sym.area.available", areaAvailable
    AssertRelative stats, "sym.area", props.Area, analyticalArea, 0.003
    AssertClose stats, "sym.cx", props.CentroidX, 0#, 0.05
    AssertClose stats, "sym.cy", props.CentroidY, 0#, 0.05
    AssertClose stats, "sym.Ixy", props.Ixyc, 0#, 0.000001 * props.Area
    AssertClose stats, "sym.principal.angle", props.PrincipalAngleRad, 0#, 0.000001
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestAsymmetricRadii(ByRef stats As TTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 300#, 180#, 10#, 35#, 20#, 50#

    Dim props25 As CSectionPropertiesCalculator
    Dim props125 As CSectionPropertiesCalculator
    Set props25 = MeshProps(geom, 25#, 25#)
    Set props125 = MeshProps(geom, 12.5, 12.5)

    AssertTrue stats, "asym.area.positive", props125.Area > 0#
    AssertTrue stats, "asym.centroid.inside", geom.ContainsPoint(props125.CentroidX, props125.CentroidY)
    AssertTrue stats, "asym.angle.valid", Abs(props125.PrincipalAngleRad) <= GEOM_PI / 2#
    AssertRelative stats, "asym.area.convergence", props125.Area, props25.Area, 0.08
    AssertRelative stats, "asym.Ix.convergence", props125.Ixc, props25.Ixc, 0.15
    AssertTrue stats, "asym.Ixy.can_be_nonzero", Abs(props125.Ixyc) > GEOM_TOLERANCE
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestInvalidData(ByRef stats As TTestStats)
    AssertInvalid stats, "invalid.width", -100#, 100#, 0#, 0#, 0#, 0#
    AssertInvalid stats, "invalid.height", 100#, 0#, 0#, 0#, 0#, 0#
    AssertInvalid stats, "invalid.negative.radius", 100#, 100#, -1#, 0#, 0#, 0#
    AssertInvalid stats, "invalid.large.radius", 100#, 100#, 101#, 0#, 0#, 0#
    AssertInvalid stats, "invalid.neighbor.radii", 100#, 100#, 60#, 60#, 0#, 0#

    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 100#, 100#, 0#, 0#, 0#, 0#
    AssertBuildError stats, "invalid.step.zero", geom, 0#, 10#
    AssertBuildError stats, "invalid.step.negative", geom, 10#, -10#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestMeshConvergence(ByRef stats As TTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 400#, 220#, 35#, 35#, 35#, 35#

    Dim p50 As CSectionPropertiesCalculator
    Dim p25 As CSectionPropertiesCalculator
    Dim p125 As CSectionPropertiesCalculator

    Set p50 = MeshProps(geom, 50#, 50#)
    Set p25 = MeshProps(geom, 25#, 25#)
    Set p125 = MeshProps(geom, 12.5, 12.5)

    AppendLine stats, "CONVERGENCE: step=50 area=" & FormatNumberInvariant(p50.Area) & _
        "; cx=" & FormatNumberInvariant(p50.CentroidX) & "; cy=" & FormatNumberInvariant(p50.CentroidY) & _
        "; Ix=" & FormatNumberInvariant(p50.Ixc) & "; Iy=" & FormatNumberInvariant(p50.Iyc) & _
        "; Ixy=" & FormatNumberInvariant(p50.Ixyc) & "; I1=" & FormatNumberInvariant(p50.PrincipalI1) & _
        "; I2=" & FormatNumberInvariant(p50.PrincipalI2)
    AppendLine stats, "CONVERGENCE: step=25 area=" & FormatNumberInvariant(p25.Area) & _
        "; cx=" & FormatNumberInvariant(p25.CentroidX) & "; cy=" & FormatNumberInvariant(p25.CentroidY) & _
        "; Ix=" & FormatNumberInvariant(p25.Ixc) & "; Iy=" & FormatNumberInvariant(p25.Iyc) & _
        "; Ixy=" & FormatNumberInvariant(p25.Ixyc) & "; I1=" & FormatNumberInvariant(p25.PrincipalI1) & _
        "; I2=" & FormatNumberInvariant(p25.PrincipalI2)
    AppendLine stats, "CONVERGENCE: step=12.5 area=" & FormatNumberInvariant(p125.Area) & _
        "; cx=" & FormatNumberInvariant(p125.CentroidX) & "; cy=" & FormatNumberInvariant(p125.CentroidY) & _
        "; Ix=" & FormatNumberInvariant(p125.Ixc) & "; Iy=" & FormatNumberInvariant(p125.Iyc) & _
        "; Ixy=" & FormatNumberInvariant(p125.Ixyc) & "; I1=" & FormatNumberInvariant(p125.PrincipalI1) & _
        "; I2=" & FormatNumberInvariant(p125.PrincipalI2)

    AssertRelative stats, "conv.area.25.12", p125.Area, p25.Area, 0.03
    AssertRelative stats, "conv.Ix.25.12", p125.Ixc, p25.Ixc, 0.05
    AssertRelative stats, "conv.Iy.25.12", p125.Iyc, p25.Iyc, 0.05
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestPerformance(ByRef stats As TTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 800#, 500#, 40#, 40#, 40#, 40#

    Dim builder As CFiberMeshBuilder
    Set builder = New CFiberMeshBuilder
    builder.BuildMesh geom, 10#, 10#, 1

    Dim calc As CSectionPropertiesCalculator
    Set calc = New CSectionPropertiesCalculator
    calc.CalculateConcrete BuildGeneratedSectionModel(builder, Nothing)

    AppendLine stats, "PERFORMANCE: section=800x500; step=10; fibers=" & CStr(builder.FiberCount) & _
        "; buildSec=" & FormatNumberInvariant(builder.BuildSeconds) & _
        "; propsSec=" & FormatNumberInvariant(calc.LastCalculationSeconds) & _
        "; area=" & FormatNumberInvariant(calc.Area)
    AssertTrue stats, "perf.fibers.positive", builder.FiberCount > 0
End Sub

Private Function MeshProps(ByVal geom As ISectionGeometry, ByVal stepX As Double, ByVal stepY As Double) As CSectionPropertiesCalculator
    On Error GoTo Failed

    Dim builder As CFiberMeshBuilder
    Set builder = New CFiberMeshBuilder
    builder.BuildMesh geom, stepX, stepY, 1
    If builder.FiberCount <= 0 Then
        Err.Raise vbObjectError + 2310, "MeshProps", "Mesh is empty; stepX=" & CStr(stepX) & "; stepY=" & CStr(stepY)
    End If

    Dim calc As CSectionPropertiesCalculator
    Set calc = New CSectionPropertiesCalculator
    calc.CalculateConcrete BuildGeneratedSectionModel(builder, Nothing)
    Set MeshProps = calc
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

Private Function MeshHasSmallFibers(ByVal mesh As CFiberMeshBuilder, ByVal baseStep As Double) As Boolean
    Dim i As Long
    For i = 1 To mesh.FiberCount
        If mesh.FiberWidth(i) < baseStep Or mesh.FiberHeight(i) < baseStep Then
            MeshHasSmallFibers = True
            Exit Function
        End If
    Next i
End Function

Private Function FirstSmallFiberWidth(ByVal mesh As CFiberMeshBuilder, ByVal baseStep As Double) As Double
    Dim i As Long
    For i = 1 To mesh.FiberCount
        If mesh.FiberWidth(i) < baseStep Then
            FirstSmallFiberWidth = mesh.FiberWidth(i)
            Exit Function
        End If
    Next i
    FirstSmallFiberWidth = 0#
End Function

Private Sub AssertInvalid(ByRef stats As TTestStats, ByVal name As String, _
        ByVal width As Double, ByVal height As Double, ByVal rtl As Double, ByVal rtr As Double, _
        ByVal rbr As Double, ByVal rbl As Double)

    Dim geom As CGeometryRoundedRectangle
    Dim message As String
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize width, height, rtl, rtr, rbr, rbl
    AssertTrue stats, name, Not geom.IsValid(message)
End Sub

Private Sub AssertBuildError(ByRef stats As TTestStats, ByVal name As String, _
        ByVal geom As CGeometryRoundedRectangle, ByVal stepX As Double, ByVal stepY As Double)

    On Error GoTo GotError
    Dim builder As CFiberMeshBuilder
    Set builder = New CFiberMeshBuilder
    builder.BuildMesh geom, stepX, stepY, 1
    On Error GoTo 0
    AssertTrue stats, name, False
    Exit Sub

GotError:
    On Error GoTo 0
    AssertTrue stats, name, True
End Sub

Private Sub AssertCircleRebarError(ByRef stats As TTestStats, ByVal name As String, _
        ByVal diameter As Double, ByVal axisDistance As Double, ByVal barCount As Long, ByVal barDiameter As Double, _
        Optional ByVal row2Diameter As Double = 0#, Optional ByVal row3Diameter As Double = 0#, _
        Optional ByVal row2Location As String = "Stacked", Optional ByVal row3Location As String = "Stacked")

    On Error GoTo GotError
    Dim builder As CCircleRebarLayoutBuilder
    Set builder = New CCircleRebarLayoutBuilder
    Dim layout As CRebarLayout
    Set layout = builder.Build(diameter, 0#, 0#, axisDistance, barCount, barDiameter, "A400", _
        row2Diameter, row3Diameter, row2Location, row3Location)
    On Error GoTo 0
    AssertTrue stats, name, False
    Exit Sub

GotError:
    On Error GoTo 0
    AssertTrue stats, name, True
End Sub

Private Sub AssertLShapeRebarError(ByRef stats As TTestStats, ByVal name As String, _
        ByVal lowerWidth As Double, ByVal upperHeight As Double, ByVal upperWidth As Double, ByVal lowerHeight As Double, _
        ByVal barCount As Long, ByVal axisDistance As Double, ByVal barDiameter As Double, _
        Optional ByVal startOffset As Double = 100#, Optional ByVal endOffset As Double = 100#)

    On Error GoTo GotError
    Dim builder As CLShapeRebarLayoutBuilder
    Set builder = New CLShapeRebarLayoutBuilder
    Dim layout As CRebarLayout
    Set layout = builder.Build(upperWidth, upperHeight, lowerWidth, lowerHeight, 0#, 0#, _
        Array(axisDistance, axisDistance, barDiameter, barDiameter, barCount, 0, startOffset, endOffset, startOffset, endOffset), _
        Array(axisDistance, axisDistance, barDiameter, barDiameter, 0, 0, startOffset, endOffset, startOffset, endOffset), _
        Array(axisDistance, axisDistance, barDiameter, barDiameter, 0, 0, startOffset, endOffset, startOffset, endOffset), _
        Array(axisDistance, axisDistance, barDiameter, barDiameter, 0, 0, startOffset, endOffset, startOffset, endOffset), _
        "A400")
    On Error GoTo 0
    AssertTrue stats, name, False
    Exit Sub

GotError:
    On Error GoTo 0
    AssertTrue stats, name, True
End Sub

Private Function EmptyFaceSettings() As Variant
    EmptyFaceSettings = Array(40#, 40#, 20#, 20#, 0, 0, 40#, 40#, 40#, 40#, 0#, 0#, 0#, 0#, "Stacked", "Stacked")
End Function

Private Function HasAnnotationGroup(ByVal layout As CRebarLayout, ByVal groupName As String) As Boolean
    Dim i As Long
    For i = 1 To layout.AnnotationCount
        If StrComp(layout.AnnotationGroupName(i), groupName, vbTextCompare) = 0 Then
            HasAnnotationGroup = True
            Exit Function
        End If
    Next i
End Function

Private Function HasSectionAnnotation(ByVal model As CSectionModel, ByVal annotationType As String, _
        ByVal annotationID As String) As Boolean
    HasSectionAnnotation = (FindSectionAnnotationIndex(model, annotationType, annotationID) > 0)
End Function

Private Function FindSectionAnnotationIndex(ByVal model As CSectionModel, ByVal annotationType As String, _
        ByVal annotationID As String) As Long
    If model Is Nothing Then Exit Function

    Dim annotations As CSectionAnnotations
    Set annotations = model.Annotations

    Dim i As Long
    For i = 1 To annotations.Count
        If StrComp(annotations.AnnotationType(i), annotationType, vbTextCompare) = 0 And _
                StrComp(annotations.AnnotationID(i), annotationID, vbTextCompare) = 0 Then
            FindSectionAnnotationIndex = i
            Exit Function
        End If
    Next i
End Function

Private Sub AssertLShapeRebarRowsError(ByRef stats As TTestStats, ByVal name As String, ByVal h1Settings As Variant)
    On Error GoTo Expected
    Dim builder As CLShapeRebarLayoutBuilder
    Set builder = New CLShapeRebarLayoutBuilder
    builder.Build 250#, 550#, 600#, 250#, 0#, 0#, h1Settings, _
        EmptyFaceSettings(), EmptyFaceSettings(), EmptyFaceSettings(), "A400"
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: " & name & "; expected validation error"
    Exit Sub
Expected:
    stats.Passed = stats.Passed + 1
    AppendLine stats, "OK: " & name
End Sub

Private Function CountBarsWithRow(ByVal layout As CRebarLayout, ByVal rowToken As String) As Long
    Dim i As Long
    For i = 1 To layout.Count
        If InStr(1, layout.BarID(i), rowToken, vbTextCompare) > 0 Then CountBarsWithRow = CountBarsWithRow + 1
    Next i
End Function

Private Sub AssertTrue(ByRef stats As TTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name
    End If
End Sub

Private Sub AssertClose(ByRef stats As TTestStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal tolerance As Double)

    Dim diff As Double
    diff = Abs(actual - expected)
    If diff <= tolerance Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected) & "; absDiff=" & FormatNumberInvariant(diff)
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected) & "; absDiff=" & FormatNumberInvariant(diff)
    End If
End Sub

Private Sub AssertRelative(ByRef stats As TTestStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal relTolerance As Double)

    Dim relDiff As Double
    If Abs(expected) <= GEOM_TOLERANCE Then
        relDiff = Abs(actual - expected)
    Else
        relDiff = Abs((actual - expected) / expected)
    End If

    If relDiff <= relTolerance Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected) & "; relDiff=" & FormatNumberInvariant(relDiff)
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected) & "; relDiff=" & FormatNumberInvariant(relDiff)
    End If
End Sub

Private Sub AppendLine(ByRef stats As TTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function
















