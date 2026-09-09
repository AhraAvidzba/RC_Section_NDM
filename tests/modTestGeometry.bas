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
    TestCircleCoreDistance stats
    TestCirclePrincipalAxesStableOnCoarseMesh stats
    TestCircleInvalidData stats
    TestCircleAutoRebarLayout stats
    TestLShapeGeometry stats
    TestLShapePrincipalAxesAndCoreDistances stats
    TestConcreteCoverUsesLocalContour stats
    TestLShapeAutoRebarLayout stats
    TestLShapeSeparateLineOffsets stats
    TestLShapeAdditionalRebarRows stats
    TestSectionModelFromGeneratedGeometry stats
    TestRebarAnnotationAnchors stats
    TestAutoCADImporterBuildsSectionModel stats
    TestAutoCADImporterRotatedRectangleBounds stats
    TestAutoCADImporterPrincipalInertiaBounds stats
    TestAutoCADImporterSquareRegionReadsEdgeRotation stats
    TestAutoCADImporterAreaSquareFallback stats
    TestAutoCADImporterAreaSquareFallbackAverageRotation stats
    TestAutoCADImporterInvalidInertiaSquareFallback stats
    TestSectionModelEquivalentRectangleFromRealInertia stats
    TestSectionModelBoundaryDoesNotReplaceRealInertia stats
    TestSectionModelFallbackInertiaSquare stats
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
    Set model = importer.BuildFromRegionArrays(concreteRegions, rebarRegions, "Rebar", 0.000001)

    AssertTrue stats, "autocad.import.source", model.SourceType = "AutoCADImport"
    AssertTrue stats, "autocad.import.concrete.count", model.ConcreteCount = 1
    AssertTrue stats, "autocad.import.rebar.count", model.RebarCount = 1
    AssertTrue stats, "autocad.import.concrete.id", model.ConcreteID(1) = "C1"
    AssertClose stats, "autocad.import.concrete.x", model.ConcreteX(1), 10#, 0.000001
    AssertClose stats, "autocad.import.concrete.y", model.ConcreteY(1), 20#, 0.000001
    AssertClose stats, "autocad.import.concrete.width", model.ConcreteWidth(1), 6#, 0.000001
    AssertClose stats, "autocad.import.concrete.height", model.ConcreteHeight(1), 3#, 0.000001
    AssertClose stats, "autocad.import.concrete.rotation", model.ConcreteRotation(1), 0#, 0.000001
    AssertClose stats, "autocad.import.concrete.localIx", model.ConcreteLocalIx(1), 18# * 3# * 3# / 12#, 0.000001
    AssertClose stats, "autocad.import.concrete.localIy", model.ConcreteLocalIy(1), 18# * 6# * 6# / 12#, 0.000001
    AssertClose stats, "autocad.import.rebar.diameter", model.RebarDiameter(1), 20#, 0.000001
    AssertTrue stats, "autocad.import.rebar.marker", model.RebarSteelClass(1) = "Rebar"
    AssertTrue stats, "autocad.import.rebar.comment", _
        InStr(1, model.RebarComment(1), "по площади", vbTextCompare) > 0

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    Dim minProjection As Double
    Dim maxProjection As Double
    props.CalculateProjection model, 1#, 0#, False, minProjection, maxProjection
    AssertClose stats, "autocad.import.concrete.xDepth", maxProjection - minProjection, 6#, 0.000001
    props.CalculateProjection model, 0#, 1#, False, minProjection, maxProjection
    AssertClose stats, "autocad.import.concrete.yDepth", maxProjection - minProjection, 3#, 0.000001
End Sub

' Проверяет, что импортированный повернутый прямоугольный элемент хранит
' реальные Width/Height/Rotation, а расчетные габариты берутся по этой форме
' без fallback-а по площади.
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
    Set model = importer.BuildFromRegionArrays(concreteRegions, rebarRegions, "Rebar", 0.000001)

    AssertTrue stats, "autocad.import.rotated.shape", model.ConcreteShapeType(1) = "Rectangle"
    AssertClose stats, "autocad.import.rotated.width", model.ConcreteWidth(1), widthValue, 0.000001
    AssertClose stats, "autocad.import.rotated.height", model.ConcreteHeight(1), heightValue, 0.000001
    AssertClose stats, "autocad.import.rotated.rotation", model.ConcreteRotation(1), angle, 0.000001
    Dim ixLocal As Double
    Dim iyLocal As Double
    Dim c As Double
    Dim s As Double
    ixLocal = concreteRegions(1, 1) * heightValue * heightValue / 12#
    iyLocal = concreteRegions(1, 1) * widthValue * widthValue / 12#
    c = Cos(angle)
    s = Sin(angle)
    AssertClose stats, "autocad.import.rotated.localIx", model.ConcreteLocalIx(1), _
        ixLocal * c * c + iyLocal * s * s, 0.000001
    AssertClose stats, "autocad.import.rotated.localIy", model.ConcreteLocalIy(1), _
        ixLocal * s * s + iyLocal * c * c, 0.000001
    AssertClose stats, "autocad.import.rotated.localIxy", model.ConcreteLocalIxy(1), _
        (iyLocal - ixLocal) * s * c, 0.000001

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    Dim minProjection As Double
    Dim maxProjection As Double
    props.CalculateProjection model, Cos(angle), Sin(angle), False, minProjection, maxProjection
    AssertClose stats, "autocad.import.rotated.uDepth", maxProjection - minProjection, widthValue, 0.000001
    props.CalculateProjection model, -Sin(angle), Cos(angle), False, minProjection, maxProjection
    AssertClose stats, "autocad.import.rotated.vDepth", maxProjection - minProjection, heightValue, 0.000001
End Sub

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
    RotatedLocalInertia inertiaU, inertiaV, angle, localIx, localIy, localIxy

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
    Set model = importer.BuildFromRegionArrays(concreteRegions, rebarRegions, "Rebar", 0.000001)

    AssertTrue stats, "autocad.import.principal.shape", model.ConcreteShapeType(1) = "Rectangle"
    AssertClose stats, "autocad.import.principal.localIx", model.ConcreteLocalIx(1), localIx, 0.000001
    AssertClose stats, "autocad.import.principal.localIy", model.ConcreteLocalIy(1), localIy, 0.000001
    AssertClose stats, "autocad.import.principal.localIxy", model.ConcreteLocalIxy(1), localIxy, 0.000001
    AssertClose stats, "autocad.import.principal.areaShell", _
        model.ConcreteWidth(1) * model.ConcreteHeight(1), area, 0.000001
    AssertClose stats, "autocad.import.principal.noExplode", importer.DebugEdgeProbeCount, 0#, 0.000001

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    Dim minProjection As Double
    Dim maxProjection As Double
    props.CalculateProjection model, Cos(angle), Sin(angle), False, minProjection, maxProjection
    AssertClose stats, "autocad.import.principal.uDepth", maxProjection - minProjection, expectedWidth, 0.000001
    props.CalculateProjection model, -Sin(angle), Cos(angle), False, minProjection, maxProjection
    AssertClose stats, "autocad.import.principal.vDepth", maxProjection - minProjection, expectedHeight, 0.000001
End Sub

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
    RotatedLocalInertia rectArea * rectHeight * rectHeight / 12#, _
        rectArea * rectWidth * rectWidth / 12#, rectAngle, rectIx, rectIy, rectIxy
    modelSpace.Add FakeRegion(rectArea, 0#, 0#, rectIx, rectIy, rectIxy, _
        "Concrete", "C_RECT", rectAngle, True)

    Dim squareArea As Double
    Dim squareAngle As Double
    squareArea = 2500#
    squareAngle = GEOM_PI / 5#
    modelSpace.Add FakeRegion(squareArea, 120#, 0#, squareArea * squareArea / 12#, _
        squareArea * squareArea / 12#, 0#, "Concrete", "C_SQ", squareAngle, True)

    modelSpace.Add FakeRegion(GEOM_PI * 12# * 12# / 4#, 0#, -60#, 1#, 1#, 0#, _
        "Reinf", "R1", 0#, False)

    Dim importer As CAutoCADSectionModelImporter
    Set importer = New CAutoCADSectionModelImporter
    Dim model As CSectionModel
    Set model = importer.ImportFromModelSpace(modelSpace, "Concrete", "Reinf", "Rebar", 0.000001)

    AssertTrue stats, "autocad.import.edgeProbe.rect.shape", model.ConcreteShapeType(1) = "Rectangle"
    AssertClose stats, "autocad.import.edgeProbe.rect.noExplode", importer.DebugEdgeProbeCount, 1#, 0.000001
    AssertTrue stats, "autocad.import.edgeProbe.square.shape", model.ConcreteShapeType(2) = "Equivalent square"
    AssertClose stats, "autocad.import.edgeProbe.square.width", model.ConcreteWidth(2), Sqr(squareArea), 0.000001
    AssertClose stats, "autocad.import.edgeProbe.square.height", model.ConcreteHeight(2), Sqr(squareArea), 0.000001
    AssertClose stats, "autocad.import.edgeProbe.square.rotation", model.ConcreteRotation(2), squareAngle, 0.000001
End Sub

' Проверяет последний fallback для Region без локальных инерций: габарит
' берется как квадрат той же площади, а не как круговой радиус.
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
    Set model = importer.BuildFromRegionArrays(concreteRegions, rebarRegions, "Rebar", 0.000001)

    AssertTrue stats, "autocad.import.squareFallback.shape", model.ConcreteShapeType(1) = "Equivalent square"
    AssertClose stats, "autocad.import.squareFallback.width", model.ConcreteWidth(1), Sqr(model.ConcreteArea(1)), 0.000001
    AssertClose stats, "autocad.import.squareFallback.height", model.ConcreteHeight(1), Sqr(model.ConcreteArea(1)), 0.000001
    AssertClose stats, "autocad.import.squareFallback.localIx", model.ConcreteLocalIx(1), _
        model.ConcreteArea(1) * model.ConcreteArea(1) / 12#, 0.000001

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    Dim minProjection As Double
    Dim maxProjection As Double
    props.CalculateProjection model, 1#, 0#, False, minProjection, maxProjection
    AssertClose stats, "autocad.import.squareFallback.xDepth", maxProjection - minProjection, Sqr(model.ConcreteArea(1)), 0.000001
    props.CalculateProjection model, 1#, 1#, False, minProjection, maxProjection
    AssertClose stats, "autocad.import.squareFallback.diagonalDepth", maxProjection - minProjection, _
        Sqr(model.ConcreteArea(1)) * Sqr(2#), 0.000001
End Sub

' Проверяет, что fallback-квадрат по площади получает средний угол
' распознанной импортированной сетки и поэтому одинаково работает в
' расчетных проекциях, Excel-схеме и AutoCAD export.
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
    Set model = importer.BuildFromRegionArrays(concreteRegions, rebarRegions, "Rebar", 0.000001)

    AssertTrue stats, "autocad.import.squareFallbackRotation.shape", model.ConcreteShapeType(2) = "Equivalent square"
    AssertClose stats, "autocad.import.squareFallbackRotation.angle", model.ConcreteRotation(2), angle, 0.000001

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    Dim minProjection As Double
    Dim maxProjection As Double
    props.CalculateProjection model, Cos(angle), Sin(angle), False, minProjection, maxProjection
    Dim expectedProjectionDepth As Double
    expectedProjectionDepth = 200# * Cos(angle) + Sqr(model.ConcreteArea(2)) / 2# + widthValue / 2#
    AssertClose stats, "autocad.import.squareFallbackRotation.projectedSide", _
        maxProjection - minProjection, expectedProjectionDepth, 0.000001
End Sub

' Проверяет, что некорректная матрица Ix/Iy/Ixy импортированного Region не
' ломает габариты устойчивости: если A/I нельзя трактовать физически, общий
' геометрический fallback переходит на квадрат той же площади.
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
    Set model = importer.BuildFromRegionArrays(concreteRegions, rebarRegions, "Rebar", 0.000001)

    AssertTrue stats, "autocad.import.badInertia.shape", model.ConcreteShapeType(1) = "Equivalent square"
    AssertClose stats, "autocad.import.badInertia.localIxy", model.ConcreteLocalIxy(1), 0#, 0.000001

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    Dim minProjection As Double
    Dim maxProjection As Double
    props.CalculateProjection model, 1#, 0#, False, minProjection, maxProjection
    AssertClose stats, "autocad.import.badInertia.xDepth", maxProjection - minProjection, Sqr(model.ConcreteArea(1)), 0.000001
End Sub

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
    RotatedLocalInertia i2, i1, angle, localIx, localIy, localIxy

    Dim model As CSectionModel
    Set model = New CSectionModel
    model.AddConcreteElement 0#, 0#, area, 1, "arbitrary", "H1", "Region", _
        0#, 0#, 0#, vbNullString, localIx, localIy, localIxy

    AssertTrue stats, "model.boundary.equivalentRectangle.status", _
        model.ConcreteShapeType(1) = "Equivalent rectangle"
    AssertClose stats, "model.boundary.equivalentRectangle.area", _
        model.ConcreteWidth(1) * model.ConcreteHeight(1), area, 0.000001
    AssertClose stats, "model.boundary.equivalentRectangle.ratio", _
        (model.ConcreteHeight(1) / model.ConcreteWidth(1)) ^ 2, i1 / i2, 0.000001
    AssertClose stats, "model.boundary.equivalentRectangle.realIx", model.ConcreteLocalIx(1), localIx, 0.000001
    AssertClose stats, "model.boundary.equivalentRectangle.realIy", model.ConcreteLocalIy(1), localIy, 0.000001
    AssertClose stats, "model.boundary.equivalentRectangle.realIxy", model.ConcreteLocalIxy(1), localIxy, 0.000001
End Sub

' Проверяет главный контракт: оболочка элемента не подменяет реальные
' локальные инерции, по которым считаются свойства бетонного сечения.
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

    AssertClose stats, "model.boundary.realIx.priorityA", propsA.Ixc, realIx, 0.000001
    AssertClose stats, "model.boundary.realIy.priorityA", propsA.Iyc, realIy, 0.000001
    AssertClose stats, "model.boundary.realIxy.priorityA", propsA.Ixyc, realIxy, 0.000001
    AssertClose stats, "model.boundary.realIx.independent", propsB.Ixc, propsA.Ixc, 0.000001
    AssertClose stats, "model.boundary.realIy.independent", propsB.Iyc, propsA.Iyc, 0.000001
    AssertClose stats, "model.boundary.realIxy.independent", propsB.Ixyc, propsA.Ixyc, 0.000001
End Sub

' Проверяет последний расчетный fallback для элемента без сохраненных
' инерций: CSectionModel возвращает квадрат по площади и валидные I.
Private Sub TestSectionModelFallbackInertiaSquare(ByRef stats As TTestStats)
    Dim area As Double
    area = 3600#

    Dim model As CSectionModel
    Set model = New CSectionModel
    model.AddConcreteElement 0#, 0#, area, 1, "noI", "noI", "Region"

    AssertTrue stats, "model.fallback.square.status", model.ConcreteShapeType(1) = "Equivalent square"
    AssertClose stats, "model.fallback.square.width", model.ConcreteWidth(1), 60#, 0.000001
    AssertClose stats, "model.fallback.square.localIx", model.ConcreteLocalIx(1), area * area / 12#, 0.000001
    AssertClose stats, "model.fallback.square.localIy", model.ConcreteLocalIy(1), area * area / 12#, 0.000001
    AssertClose stats, "model.fallback.square.localIxy", model.ConcreteLocalIxy(1), 0#, 0.000001
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

' Проверяет, что PrincipalAngle задает физическую главную ось 1, а ядровые
' расстояния главных плоскостей считаются через ту же нормаль, что и
' устойчивость. Тест защищает Г-сечение, где ошибка знака угла сразу дает
' заметный ненулевой I12 после поворота.
Private Sub TestLShapePrincipalAxesAndCoreDistances(ByRef stats As TTestStats)
    Dim geom As CGeometryLShape
    Set geom = New CGeometryLShape
    geom.Initialize 250#, 550#, 600#, 250#

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 25#, 25#, 1, 2

    Dim concreteSection As CSectionModel
    Set concreteSection = BuildGeneratedSectionModel(mesh, Nothing, "LShapePrincipalConcrete")

    Dim concreteProps As CSectionPropertiesCalculator
    Set concreteProps = New CSectionPropertiesCalculator
    concreteProps.CalculateConcrete concreteSection
    AssertPrincipalAxesConsistent stats, "lshape.concrete", concreteSection, concreteProps

    Dim rebarBuilder As CLShapeRebarLayoutBuilder
    Set rebarBuilder = New CLShapeRebarLayoutBuilder
    Dim rebars As CRebarLayout
    Set rebars = rebarBuilder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        Array(40#, 40#, 32#, 32#, 5, 5, 80#, 80#, 80#, 80#), _
        Array(40#, 40#, 32#, 32#, 2, 2, 80#, 80#, 80#, 80#), _
        Array(40#, 40#, 32#, 32#, 2, 2, 80#, 80#, 80#, 80#), _
        Array(40#, 40#, 32#, 32#, 5, 5, 80#, 80#, 80#, 80#), _
        "Rebar")

    Dim transformedSection As CSectionModel
    Set transformedSection = BuildGeneratedSectionModel(mesh, rebars, "LShapePrincipalTransformed")

    Dim transformedProps As CSectionPropertiesCalculator
    Set transformedProps = New CSectionPropertiesCalculator
    transformedProps.CalculateTransformedByModuli transformedSection, 32500#, 200000#
    AssertPrincipalAxesConsistent stats, "lshape.transformed", transformedSection, transformedProps
End Sub

' Проверяет общие инварианты главных осей для любого представления сечения.
' Ось 1 должна давать I1, ось 2 - I2, а I12 в повернутой системе должен
' исчезать. CoreDistance главной плоскости сверяется с прямым расчетом по
' соответствующей нормали к оси изгиба.
Private Sub AssertPrincipalAxesConsistent(ByRef stats As TTestStats, ByVal prefix As String, _
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

    AssertRelative stats, prefix & ".axis1.inertia", inertia1, props.PrincipalI1, 0.0000001
    AssertRelative stats, prefix & ".axis2.inertia", inertia2, props.PrincipalI2, 0.0000001

    Dim c As Double
    Dim s As Double
    Dim rotatedIxy As Double
    c = Cos(props.PrincipalAngleRad)
    s = Sin(props.PrincipalAngleRad)
    rotatedIxy = (props.Ixc - props.Iyc) * s * c + props.Ixyc * (c * c - s * s)
    AssertTrue stats, prefix & ".principal.I12.zero", _
        Abs(rotatedIxy) <= GeomMax(props.PrincipalI1, 1#) * 0.0000001

    Dim normal1X As Double
    Dim normal1Y As Double
    props.PrincipalPlaneNormalDirection 1, normal1X, normal1Y
    AssertClose stats, prefix & ".core.plane1.plus", _
        props.PrincipalPlaneCoreDistance(section, 1, True, False), _
        props.CoreDistanceAlong(section, normal1X, normal1Y, False), 0.000001
    AssertClose stats, prefix & ".core.plane1.minus", _
        props.PrincipalPlaneCoreDistance(section, 1, False, False), _
        props.CoreDistanceAlong(section, -normal1X, -normal1Y, False), 0.000001

    Dim normal2X As Double
    Dim normal2Y As Double
    props.PrincipalPlaneNormalDirection 2, normal2X, normal2Y
    AssertClose stats, prefix & ".core.plane2.plus", _
        props.PrincipalPlaneCoreDistance(section, 2, True, False), _
        props.CoreDistanceAlong(section, normal2X, normal2Y, False), 0.000001
    AssertClose stats, prefix & ".core.plane2.minus", _
        props.PrincipalPlaneCoreDistance(section, 2, False, False), _
        props.CoreDistanceAlong(section, -normal2X, -normal2Y, False), 0.000001
End Sub

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

    Dim localCover As Double
    localCover = props.ConcreteCoverFromPointAlongDirection(section, 0#, 60#, 0#, 1#)
    AssertClose stats, "geometry.cover.localContour", localCover, 40#, 0.000001

    Dim minS As Double
    Dim maxS As Double
    props.CalculateProjection section, 0#, 1#, False, minS, maxS
    AssertClose stats, "geometry.cover.globalWouldBeWrong", maxS - 60#, 340#, 0.000001
    AssertTrue stats, "geometry.cover.localLessThanGlobal", localCover < maxS - 60#
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

' Проверяет ядровое расстояние на круге: для сплошного круга r = R/4.
' Метод CoreDistanceAlong затем используется устойчивостью СП 35 для ветви e/r.
Private Sub TestCircleCoreDistance(ByRef stats As TTestStats)
    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter 300#

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 2.5, 2.5, 1

    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, Nothing)

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete section

    AssertRelative stats, "circle.core.xPositive", props.CoreDistanceAlong(section, 1#, 0#, False), 37.5, 0.02
    AssertRelative stats, "circle.core.yNegative", props.CoreDistanceAlong(section, 0#, -1#, False), 37.5, 0.02
End Sub

' Проверяет, что грубая, но симметричная сетка круга не разворачивает главные
' оси из-за микроскопического численного Ixy. Для устойчивости это важно:
' чистый пользовательский Mx не должен попадать во вторую плоскость только из-за
' дискретизационного шума.
Private Sub TestCirclePrincipalAxesStableOnCoarseMesh(ByRef stats As TTestStats)
    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter 500#

    Dim props As CSectionPropertiesCalculator
    Set props = MeshProps(geom, 20#, 20#)

    AssertClose stats, "circle.coarse.principal.angle", props.PrincipalAngleRad, 0#, 0.000000001
    AssertTrue stats, "circle.coarse.IxIy.nearlyEqual", _
        Abs(props.Ixc - props.Iyc) / GeomMax(Abs(props.Ixc), Abs(props.Iyc)) < 0.000001
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
    AssertClose stats, "sym.principal.angle", props.PrincipalAngleRad, GEOM_PI / 2#, 0.000001
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

' Формирует локальные Ix/Iy/Ixy после поворота главных осей элемента.
' minorInertia относится к оси вдоль локальной стороны width, majorInertia -
' к перпендикулярной главной оси.
Private Sub RotatedLocalInertia(ByVal minorInertia As Double, ByVal majorInertia As Double, _
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

' Создает минимальный fake AutoCAD Region для тестов live-пути importer-а.
Private Function FakeRegion(ByVal area As Double, ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal localIx As Double, ByVal localIy As Double, ByVal localIxy As Double, _
        ByVal layerName As String, ByVal handleText As String, _
        Optional ByVal edgeRotation As Double = 0#, Optional ByVal hasEdge As Boolean = False) As CFakeAcadRegion
    Dim region As CFakeAcadRegion
    Set region = New CFakeAcadRegion
    region.Initialize area, xCoord, yCoord, localIx, localIy, localIxy, _
        layerName, handleText, edgeRotation, hasEdge
    Set FakeRegion = region
End Function

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
















