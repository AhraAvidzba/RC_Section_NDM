Attribute VB_Name = "modTestMaterialDiagrams"
Option Explicit

' ==========================================================================
' Тесты построения диаграмм материалов
' ==========================================================================
' Проверяется два уровня новой архитектуры: builders формируют нормативные
' TwoLine/ThreeLine точки из параметров материала, а расчетные material-классы
' дальше только интерполируют уже готовую диаграмму.

Private Type TMaterialTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' Запускает связанный набор операций и возвращает пользователю итоговый статус выполнения.
Public Function RunMaterialDiagramTests() As String
    On Error GoTo Failed

    Dim stats As TMaterialTestStats
    Dim t0 As Double
    t0 = Timer

    TestConcreteDiagramBuilder stats
    TestSteelDiagramBuilder stats
    TestConcreteInterpolation stats
    TestSteelInterpolation stats
    TestConcreteThreePointDiagram stats
    TestSteelThreePointDiagram stats
    TestInvalidParameters stats

    AppendLine stats, "TOTAL_MATERIAL: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunMaterialDiagramTests = stats.Report
    Exit Function

Failed:
    RunMaterialDiagramTests = "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestConcreteInterpolation(ByRef stats As TMaterialTestStats)
    Dim mat As CConcreteDiagramMaterial
    Set mat = New CConcreteDiagramMaterial
    mat.Initialize -0.0015, -15.5, -0.0035, -15.5

    AssertClose stats, "conc.points.tension.zero", mat.GetStress(0.0001), 0#, 0.000000000001
    AssertClose stats, "conc.points.origin", mat.GetStress(0#), 0#, 0.000000000001
    AssertClose stats, "conc.points.firstSegment", mat.GetStress(-0.00075), -7.75, 0.000000000001
    AssertClose stats, "conc.points.secondSegment", mat.GetStress(-0.0025), -15.5, 0.000000000001
    AssertClose stats, "conc.points.secondTangent", mat.GetTangentModulus(-0.0025), 0#, 0.000000000001
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestConcreteThreePointDiagram(ByRef stats As TMaterialTestStats)
    Dim mat As CConcreteDiagramMaterial
    Set mat = New CConcreteDiagramMaterial
    mat.InitializeFromThreeCompressionPoints -0.001, -10#, -0.002, -15#, -0.003, -18#

    AssertClose stats, "conc.threePoints.firstSegment", mat.GetStress(-0.0005), -5#, 0.000000000001
    AssertClose stats, "conc.threePoints.secondSegment", mat.GetStress(-0.0015), -12.5, 0.000000000001
    AssertClose stats, "conc.threePoints.thirdSegment", mat.GetStress(-0.0025), -16.5, 0.000000000001
    AssertClose stats, "conc.threePoints.ultimate", mat.GetStress(-0.004), -18#, 0.000000000001
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestSteelThreePointDiagram(ByRef stats As TMaterialTestStats)
    Dim mat As CSteelDiagramMaterial
    Set mat = New CSteelDiagramMaterial
    mat.InitializeFromThreeTensionPoints 0.001, 200#, 0.01, 350#, 0.025, 420#

    AssertClose stats, "steel.threePoints.firstSegment", mat.GetStress(0.0005), 100#, 0.000000000001
    AssertClose stats, "steel.threePoints.secondSegment", mat.GetStress(0.0055), 275#, 0.000000000001
    AssertClose stats, "steel.threePoints.thirdSegment", mat.GetStress(0.0175), 385#, 0.000000000001
    AssertClose stats, "steel.threePoints.compression", mat.GetStress(-0.0175), -385#, 0.000000000001
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestSteelInterpolation(ByRef stats As TMaterialTestStats)
    Dim mat As CSteelDiagramMaterial
    Set mat = New CSteelDiagramMaterial
    mat.Initialize 0.00175, 350#, 0.025

    AssertClose stats, "steel.points.tensionFirstSegment", mat.GetStress(0.000875), 175#, 0.000000000001
    AssertClose stats, "steel.points.tensionSecondSegment", mat.GetStress(0.01), 350#, 0.000000000001
    AssertClose stats, "steel.points.compressionFirstSegment", mat.GetStress(-0.000875), -175#, 0.000000000001
    AssertClose stats, "steel.points.compressionSecondSegment", mat.GetStress(-0.01), -350#, 0.000000000001
    AssertClose stats, "steel.points.secondTangent", mat.GetTangentModulus(0.01), 0#, 0.000000000001
End Sub

' Проверяет бетонный builder на контрольных short-term точках СП 63:
' I ГПС без растянутого бетона для прочности и II ГПС с растянутой ветвью
' для состояния образования трещины.
Private Sub TestConcreteDiagramBuilder(ByRef stats As TMaterialTestStats)
    Dim parameters As CConcreteMaterialParameters
    Set parameters = TestConcreteParameters()

    Dim builder As CConcreteDiagramBuilder
    Set builder = New CConcreteDiagramBuilder

    Dim strength As CConcreteDiagramMaterial
    Set strength = builder.BuildMaterial(parameters, "I", "TwoLine", "Ignore")
    AssertClose stats, "conc.builder.strength.point1.eps", strength.PointStrain(1), -0.0035, 0.000000000001
    AssertClose stats, "conc.builder.strength.point1.stress", strength.PointStress(1), -15.5, 0.000000000001
    AssertClose stats, "conc.builder.strength.point2.eps", strength.PointStrain(2), -0.0015, 0.000000000001
    AssertClose stats, "conc.builder.strength.tensionZero", strength.GetStress(0.0001), 0#, 0.000000000001
    AssertClose stats, "conc.builder.strength.compressionLimit", builder.CompressionLimit, -0.0035, 0.000000000001
    AssertClose stats, "conc.builder.strength.tensionLimit", builder.TensionLimit, 0#, 0.000000000001

    Dim mcrc As CConcreteDiagramMaterial
    Set mcrc = builder.BuildMaterial(parameters, "II", "ThreeLine", "UseDiagram")
    AssertClose stats, "conc.builder.mcrc.compressionRb", mcrc.GetStress(-0.002), -22#, 0.000000000001
    AssertClose stats, "conc.builder.mcrc.tensionRbt", mcrc.GetStress(0.0001), 1.8, 0.000000000001
    AssertClose stats, "conc.builder.mcrc.epsB1", mcrc.PointStrain(3), -(0.6 * 22# / 32500#), 0.000000000001
    AssertClose stats, "conc.builder.mcrc.epsBt1", mcrc.PointStrain(5), 0.6 * 1.8 / 32500#, 0.000000000001
    AssertClose stats, "conc.builder.mcrc.tensionLimit", builder.TensionLimit, 0.00015, 0.000000000001
End Sub

' Проверяет построение диаграмм арматуры по пп. 6.2.14 и 6.2.15 СП 63.
' Для ThreeLine отдельно контролируются eps_s0, eps_s1 и eps_s,pl, потому что
' эти точки определяют участок условного предела текучести.
Private Sub TestSteelDiagramBuilder(ByRef stats As TMaterialTestStats)
    Dim parameters As CSteelMaterialParameters
    Set parameters = TestSteelParameters()

    Dim builder As CSteelDiagramBuilder
    Set builder = New CSteelDiagramBuilder

    Dim twoLine As CSteelDiagramMaterial
    Set twoLine = builder.BuildMaterial(parameters, "I", "TwoLine")
    AssertClose stats, "steel.builder.twoline.epsSc0", twoLine.PointStrain(2), -350# / 200000#, 0.000000000001
    AssertClose stats, "steel.builder.twoline.epsS0", twoLine.PointStrain(4), 350# / 200000#, 0.000000000001
    AssertClose stats, "steel.builder.twoline.limit", builder.TensionLimit, 0.025, 0.000000000001

    Dim threeLine As CSteelDiagramMaterial
    Set threeLine = builder.BuildMaterial(parameters, "I", "ThreeLine")
    Dim epsS0 As Double
    Dim epsS1 As Double
    Dim epsSPl As Double
    epsS0 = 350# / 200000# + 0.002
    epsS1 = 0.9 * 350# / 200000#
    epsSPl = 2# * epsS0 - epsS1

    AssertClose stats, "steel.builder.threeline.epsS1", threeLine.PointStrain(6), epsS1, 0.000000000001
    AssertClose stats, "steel.builder.threeline.epsS0", threeLine.PointStrain(7), epsS0, 0.000000000001
    AssertClose stats, "steel.builder.threeline.epsSPl", threeLine.PointStrain(8), epsSPl, 0.000000000001
    AssertClose stats, "steel.builder.threeline.limit", builder.TensionLimit, 0.015, 0.000000000001
    Dim slsTwoLine As CSteelDiagramMaterial
    Set slsTwoLine = builder.BuildMaterial(parameters, "II", "TwoLine")
    AssertClose stats, "steel.builder.sls.resistance", _
        slsTwoLine.GetStress(0.01), 390#, 0.000000000001
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestInvalidParameters(ByRef stats As TMaterialTestStats)
    On Error GoTo ConcreteError
    Dim concrete As CConcreteDiagramMaterial
    Set concrete = New CConcreteDiagramMaterial
    concrete.Initialize 0.0015, -15.5, -0.0035, -15.5
    AssertTrue stats, "material.invalidConcrete", False
    GoTo SteelCheck

ConcreteError:
    On Error GoTo 0
    AssertTrue stats, "material.invalidConcrete", True
    Resume SteelCheck

SteelCheck:
    On Error GoTo SteelError
    Dim steel As CSteelDiagramMaterial
    Set steel = New CSteelDiagramMaterial
    steel.Initialize 0#, 350#, 0.025
    AssertTrue stats, "material.invalidSteel", False
    Exit Sub

SteelError:
    On Error GoTo 0
    AssertTrue stats, "material.invalidSteel", True
End Sub

Private Function TestConcreteParameters() As CConcreteMaterialParameters
    Dim parameters As CConcreteMaterialParameters
    Set parameters = New CConcreteMaterialParameters
    parameters.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#
    Set TestConcreteParameters = parameters
End Function

Private Function TestSteelParameters() As CSteelMaterialParameters
    Dim parameters As CSteelMaterialParameters
    Set parameters = New CSteelMaterialParameters
    parameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#, "Ribbed"
    Set TestSteelParameters = parameters
End Function

Private Sub AssertTrue(ByRef stats As TMaterialTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name
    End If
End Sub

Private Sub AssertClose(ByRef stats As TMaterialTestStats, ByVal name As String, ByVal actual As Double, _
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

Private Sub AppendLine(ByRef stats As TMaterialTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function







