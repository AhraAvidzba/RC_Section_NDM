Attribute VB_Name = "modTestMaterialDiagrams"
Option Explicit

' ==========================================================================
' Тесты пользовательских диаграмм материалов
' ==========================================================================
' Проверяется кусочно-линейная интерполяция бетона и арматуры по точкам,
' включая режимы работы растянутого бетона.

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

    TestConcreteDiagramPoints stats
    TestSteelDiagramPoints stats
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
Private Sub TestConcreteDiagramPoints(ByRef stats As TMaterialTestStats)
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
Private Sub TestSteelDiagramPoints(ByRef stats As TMaterialTestStats)
    Dim mat As CSteelDiagramMaterial
    Set mat = New CSteelDiagramMaterial
    mat.Initialize 0.00175, 350#, 0.025

    AssertClose stats, "steel.points.tensionFirstSegment", mat.GetStress(0.000875), 175#, 0.000000000001
    AssertClose stats, "steel.points.tensionSecondSegment", mat.GetStress(0.01), 350#, 0.000000000001
    AssertClose stats, "steel.points.compressionFirstSegment", mat.GetStress(-0.000875), -175#, 0.000000000001
    AssertClose stats, "steel.points.compressionSecondSegment", mat.GetStress(-0.01), -350#, 0.000000000001
    AssertClose stats, "steel.points.secondTangent", mat.GetTangentModulus(0.01), 0#, 0.000000000001
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







