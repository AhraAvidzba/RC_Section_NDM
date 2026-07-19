Attribute VB_Name = "modTestGeometry"
Option Explicit

Private Type TTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

Public Function RunGeometryTests() As String
    On Error GoTo Failed

    Dim stats As TTestStats

    TestRectangle stats
    TestSymmetricRoundedRectangle stats
    TestAsymmetricRadii stats
    TestCircleGeometry stats
    TestCircleInvalidData stats
    TestCircleAutoRebarLayout stats
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

Private Sub TestRectangle(ByRef stats As TTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 200#, 100#, 0#, 0#, 0#, 0#

    Dim props As CGeometryPropertiesCalculator
    Set props = MeshProps(geom, 2.5, 2.5)

    AssertClose stats, "rect.area", props.Area, 200# * 100#, 0.000001
    AssertClose stats, "rect.cx", props.CentroidX, 0#, 0.000001
    AssertClose stats, "rect.cy", props.CentroidY, 0#, 0.000001
    AssertRelative stats, "rect.Ix", props.Ixc, 200# * 100# ^ 3 / 12#, 0.001
    AssertRelative stats, "rect.Iy", props.Iyc, 100# * 200# ^ 3 / 12#, 0.001
    AssertClose stats, "rect.Ixy", props.Ixyc, 0#, 0.000001
End Sub

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

    AssertCircleRebarError stats, "circle.rebar.invalid.D", 0#, 40#, 8, 20#
    AssertCircleRebarError stats, "circle.rebar.invalid.n", 300#, 40#, 1, 20#
    AssertCircleRebarError stats, "circle.rebar.invalid.ds", 300#, 40#, 8, 0#
    AssertCircleRebarError stats, "circle.rebar.invalid.as.small", 300#, 10#, 8, 20#
    AssertCircleRebarError stats, "circle.rebar.invalid.as.large", 300#, 150#, 8, 20#
End Sub

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

    Dim centerProps As CGeometryPropertiesCalculator
    Set centerProps = New CGeometryPropertiesCalculator
    centerProps.CalculateFromMesh centerMesh

    Dim subcellProps As CGeometryPropertiesCalculator
    Set subcellProps = New CGeometryPropertiesCalculator
    subcellProps.CalculateFromMesh subcellMesh

    Dim targetArea As Double
    targetArea = GEOM_PI * 95# * 95#

    AssertTrue stats, "boundary.subcell.fibers.more", subcellMesh.FiberCount > centerMesh.FiberCount
    AssertTrue stats, "boundary.subcell.area.better", Abs(subcellProps.Area - targetArea) < Abs(centerProps.Area - targetArea)
    AssertTrue stats, "boundary.subcell.has.small.fibers", MeshHasSmallFibers(subcellMesh, 40#)
    AssertClose stats, "boundary.subcell.small.width", FirstSmallFiberWidth(subcellMesh, 40#), 10#, 0.000001
End Sub

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

    Dim props As CGeometryPropertiesCalculator
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

Private Sub TestSymmetricRoundedRectangle(ByRef stats As TTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 300#, 160#, 25#, 25#, 25#, 25#

    Dim areaAvailable As Boolean
    Dim analyticalArea As Double
    analyticalArea = geom.AnalyticalArea(areaAvailable)

    Dim props As CGeometryPropertiesCalculator
    Set props = MeshProps(geom, 2.5, 2.5)

    AssertTrue stats, "sym.area.available", areaAvailable
    AssertRelative stats, "sym.area", props.Area, analyticalArea, 0.003
    AssertClose stats, "sym.cx", props.CentroidX, 0#, 0.05
    AssertClose stats, "sym.cy", props.CentroidY, 0#, 0.05
    AssertClose stats, "sym.Ixy", props.Ixyc, 0#, 0.000001 * props.Area
    AssertClose stats, "sym.principal.angle", props.PrincipalAngleRad, 0#, 0.000001
End Sub

Private Sub TestAsymmetricRadii(ByRef stats As TTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 300#, 180#, 10#, 35#, 20#, 50#

    Dim props25 As CGeometryPropertiesCalculator
    Dim props125 As CGeometryPropertiesCalculator
    Set props25 = MeshProps(geom, 25#, 25#)
    Set props125 = MeshProps(geom, 12.5, 12.5)

    AssertTrue stats, "asym.area.positive", props125.Area > 0#
    AssertTrue stats, "asym.centroid.inside", geom.ContainsPoint(props125.CentroidX, props125.CentroidY)
    AssertTrue stats, "asym.angle.valid", Abs(props125.PrincipalAngleRad) <= GEOM_PI / 2#
    AssertRelative stats, "asym.area.convergence", props125.Area, props25.Area, 0.08
    AssertRelative stats, "asym.Ix.convergence", props125.Ixc, props25.Ixc, 0.15
    AssertTrue stats, "asym.Ixy.can_be_nonzero", Abs(props125.Ixyc) > GEOM_TOLERANCE
End Sub

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

Private Sub TestMeshConvergence(ByRef stats As TTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 400#, 220#, 35#, 35#, 35#, 35#

    Dim p50 As CGeometryPropertiesCalculator
    Dim p25 As CGeometryPropertiesCalculator
    Dim p125 As CGeometryPropertiesCalculator

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

Private Sub TestPerformance(ByRef stats As TTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 800#, 500#, 40#, 40#, 40#, 40#

    Dim builder As CFiberMeshBuilder
    Set builder = New CFiberMeshBuilder
    builder.BuildMesh geom, 10#, 10#, 1

    Dim calc As CGeometryPropertiesCalculator
    Set calc = New CGeometryPropertiesCalculator
    calc.CalculateFromMesh builder

    AppendLine stats, "PERFORMANCE: section=800x500; step=10; fibers=" & CStr(builder.FiberCount) & _
        "; buildSec=" & FormatNumberInvariant(builder.BuildSeconds) & _
        "; propsSec=" & FormatNumberInvariant(calc.LastCalculationSeconds) & _
        "; area=" & FormatNumberInvariant(calc.Area)
    AssertTrue stats, "perf.fibers.positive", builder.FiberCount > 0
End Sub

Private Function MeshProps(ByVal geom As ISectionGeometry, ByVal stepX As Double, ByVal stepY As Double) As CGeometryPropertiesCalculator
    On Error GoTo Failed

    Dim builder As CFiberMeshBuilder
    Set builder = New CFiberMeshBuilder
    builder.BuildMesh geom, stepX, stepY, 1
    If builder.FiberCount <= 0 Then
        Err.Raise vbObjectError + 2310, "MeshProps", "Mesh is empty; stepX=" & CStr(stepX) & "; stepY=" & CStr(stepY)
    End If

    Dim calc As CGeometryPropertiesCalculator
    Set calc = New CGeometryPropertiesCalculator
    calc.CalculateFromMesh builder
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
        ByVal diameter As Double, ByVal axisDistance As Double, ByVal barCount As Long, ByVal barDiameter As Double)

    On Error GoTo GotError
    Dim builder As CCircleRebarLayoutBuilder
    Set builder = New CCircleRebarLayoutBuilder
    Dim layout As CRebarLayout
    Set layout = builder.Build(diameter, 0#, 0#, axisDistance, barCount, barDiameter, "A400")
    On Error GoTo 0
    AssertTrue stats, name, False
    Exit Sub

GotError:
    On Error GoTo 0
    AssertTrue stats, name, True
End Sub

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













