Attribute VB_Name = "modTestMaterialDiagrams"
Option Explicit

' ==========================================================================
' Тесты диаграмм материалов
' ==========================================================================
' Проверяется целевая упрощенная архитектура:
' - CMaterialDiagram хранит готовые точки и выполняет только интерполяцию;
' - CMaterialModelProvider выбирает расчетный режим, I/II ГПС, TwoLine/ThreeLine
'   и строит точки диаграмм из параметров бетона и арматуры.
' Старые builder-классы здесь намеренно не используются.

Private Type TMaterialTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' Запускает все проверки material-слоя и возвращает текстовый отчет для Run-AllTests.
Public Function RunMaterialDiagramTests() As String
    On Error GoTo Failed

    Dim stats As TMaterialTestStats
    Dim t0 As Double
    t0 = Timer

    TestConcreteInterpolation stats
    TestSteelInterpolation stats
    TestConcreteThreePointDiagram stats
    TestSteelThreePointDiagram stats
    TestProviderConcreteDiagrams stats
    TestProviderSteelDiagrams stats
    TestProviderStateSolutionExtension stats
    TestUserStrainParametersAffectDiagrams stats
    TestInvalidParameters stats

    AppendLine stats, "TOTAL_MATERIAL: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunMaterialDiagramTests = stats.Report
    Exit Function

Failed:
    RunMaterialDiagramTests = "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

' Проверяет сжатую бетонную ветвь: до первой точки держится плато,
' между точками работает линейная интерполяция, а растяжение выключено.
Private Sub TestConcreteInterpolation(ByRef stats As TMaterialTestStats)
    Dim mat As CMaterialDiagram
    Set mat = New CMaterialDiagram
    mat.Initialize -0.0015, -15.5, -0.0035, -15.5

    AssertClose stats, "conc.points.tension.zero", mat.GetStress(0.0001), 0#, 0.000000000001
    AssertClose stats, "conc.points.origin", mat.GetStress(0#), 0#, 0.000000000001
    AssertClose stats, "conc.points.firstSegment", mat.GetStress(-0.00075), -7.75, 0.000000000001
    AssertClose stats, "conc.points.secondSegment", mat.GetStress(-0.0025), -15.5, 0.000000000001
    AssertClose stats, "conc.points.secondTangent", mat.GetTangentModulus(-0.0025), 0#, 0.000000000001
End Sub

' Проверяет трехточечную сжатую ветвь без provider-а, чтобы отдельно закрепить
' поведение универсальной кусочно-линейной интерполяции.
Private Sub TestConcreteThreePointDiagram(ByRef stats As TMaterialTestStats)
    Dim mat As CMaterialDiagram
    Set mat = New CMaterialDiagram
    mat.InitializeFromThreeCompressionPoints -0.001, -10#, -0.002, -15#, -0.003, -18#

    AssertClose stats, "conc.threePoints.firstSegment", mat.GetStress(-0.0005), -5#, 0.000000000001
    AssertClose stats, "conc.threePoints.secondSegment", mat.GetStress(-0.0015), -12.5, 0.000000000001
    AssertClose stats, "conc.threePoints.thirdSegment", mat.GetStress(-0.0025), -16.5, 0.000000000001
    AssertClose stats, "conc.threePoints.ultimate", mat.GetStress(-0.004), -18#, 0.000000000001
End Sub

' Проверяет симметричную трехлинейную диаграмму арматуры как готовую диаграмму,
' без знания о нормативных формулах построения точек.
Private Sub TestSteelThreePointDiagram(ByRef stats As TMaterialTestStats)
    Dim mat As CMaterialDiagram
    Set mat = New CMaterialDiagram
    mat.InitializeFromThreeTensionPoints 0.001, 200#, 0.01, 350#, 0.025, 420#

    AssertClose stats, "steel.threePoints.firstSegment", mat.GetStress(0.0005), 100#, 0.000000000001
    AssertClose stats, "steel.threePoints.secondSegment", mat.GetStress(0.0055), 275#, 0.000000000001
    AssertClose stats, "steel.threePoints.thirdSegment", mat.GetStress(0.0175), 385#, 0.000000000001
    AssertClose stats, "steel.threePoints.compression", mat.GetStress(-0.0175), -385#, 0.000000000001
End Sub

' Проверяет двухлинейную диаграмму арматуры: линейный участок до R/Es
' и дальнейшее пластическое плато до пользовательской предельной деформации.
Private Sub TestSteelInterpolation(ByRef stats As TMaterialTestStats)
    Dim mat As CMaterialDiagram
    Set mat = New CMaterialDiagram
    mat.Initialize 0.00175, 350#, 0.025

    AssertClose stats, "steel.points.tensionFirstSegment", mat.GetStress(0.000875), 175#, 0.000000000001
    AssertClose stats, "steel.points.tensionSecondSegment", mat.GetStress(0.01), 350#, 0.000000000001
    AssertClose stats, "steel.points.compressionFirstSegment", mat.GetStress(-0.000875), -175#, 0.000000000001
    AssertClose stats, "steel.points.compressionSecondSegment", mat.GetStress(-0.01), -350#, 0.000000000001
    AssertClose stats, "steel.points.secondTangent", mat.GetTangentModulus(0.01), 0#, 0.000000000001
End Sub

' Проверяет, что provider выбирает нужные бетонные характеристики и режим
' растянутого бетона для Strength/Mcrc/CrackedNDS.
Private Sub TestProviderConcreteDiagrams(ByRef stats As TMaterialTestStats)
    Dim provider As CMaterialModelProvider
    Set provider = TestProvider()

    Dim strength As CMaterialDiagram
    Set strength = provider.ConcreteMaterial(cpStrength)
    AssertClose stats, "conc.provider.strength.point1.eps", strength.PointStrain(1), -0.0035, 0.000000000001
    AssertClose stats, "conc.provider.strength.point1.stress", strength.PointStress(1), -15.5, 0.000000000001
    AssertClose stats, "conc.provider.strength.point2.eps", strength.PointStrain(2), -0.0015, 0.000000000001
    AssertClose stats, "conc.provider.strength.tensionZero", strength.GetStress(0.0001), 0#, 0.000000000001
    AssertClose stats, "conc.provider.strength.compressionLimit", provider.ConcreteCompressionLimit(cpStrength), -0.0035, 0.000000000001
    AssertClose stats, "conc.provider.strength.tensionLimit", provider.ConcreteTensionLimit(cpStrength), 0#, 0.000000000001

    Dim mcrc As CMaterialDiagram
    Set mcrc = provider.ConcreteMaterial(cpMcrc)
    AssertClose stats, "conc.provider.mcrc.compressionRb", mcrc.GetStress(-0.002), -22#, 0.000000000001
    AssertClose stats, "conc.provider.mcrc.tensionRbt", mcrc.GetStress(0.0001), 1.8, 0.000000000001
    AssertClose stats, "conc.provider.mcrc.epsB1", mcrc.PointStrain(3), -(0.6 * 22# / 32500#), 0.000000000001
    AssertClose stats, "conc.provider.mcrc.epsBt1", mcrc.PointStrain(5), 0.6 * 1.8 / 32500#, 0.000000000001
    AssertClose stats, "conc.provider.mcrc.tensionLimit", provider.ConcreteTensionLimit(cpMcrc), 0.00015, 0.000000000001

    AssertClose stats, "conc.provider.cracked.tensionZero", provider.ConcreteMaterial(cpCrackedNDS).GetStress(0.0001), 0#, 0.000000000001
End Sub

' Проверяет нормативные расчетные точки арматуры, которые provider вычисляет
' из R/E и выбранной двухлинейной или трехлинейной схемы.
Private Sub TestProviderSteelDiagrams(ByRef stats As TMaterialTestStats)
    Dim provider As CMaterialModelProvider
    Set provider = TestProvider()

    Dim twoLine As CMaterialDiagram
    Set twoLine = provider.SteelMaterial(cpStrength)
    AssertClose stats, "steel.provider.twoline.epsSc0", twoLine.PointStrain(2), -350# / 200000#, 0.000000000001
    AssertClose stats, "steel.provider.twoline.epsS0", twoLine.PointStrain(4), 350# / 200000#, 0.000000000001
    AssertClose stats, "steel.provider.twoline.limit", provider.SteelTensionLimit(cpStrength), 0.025, 0.000000000001

    Dim providerThree As CMaterialModelProvider
    Set providerThree = TestProvider("TwoLine", "Ignore", "ThreeLine", "ThreeLine", "ThreeLine")

    Dim threeLine As CMaterialDiagram
    Set threeLine = providerThree.SteelMaterial(cpStrength)
    Dim epsS0 As Double
    Dim epsS1 As Double
    Dim epsSPl As Double
    epsS0 = 350# / 200000# + 0.002
    epsS1 = 0.9 * 350# / 200000#
    epsSPl = 2# * epsS0 - epsS1

    AssertClose stats, "steel.provider.threeline.epsS1", threeLine.PointStrain(6), epsS1, 0.000000000001
    AssertClose stats, "steel.provider.threeline.epsS0", threeLine.PointStrain(7), epsS0, 0.000000000001
    AssertClose stats, "steel.provider.threeline.epsSPl", threeLine.PointStrain(8), epsSPl, 0.000000000001
    AssertClose stats, "steel.provider.threeline.limit", providerThree.SteelTensionLimit(cpStrength), 0.015, 0.000000000001
    AssertClose stats, "steel.provider.sls.resistance", provider.SteelMaterial(cpCrackedNDS).GetStress(0.01), 390#, 0.000000000001
End Sub

' Проверяет, что numerical extension появляется только у StateSolution-диаграмм.
' Физические Strength/Mcrc/CrackedNDS-диаграммы остаются без технических точек,
' поэтому capacity и crack не получают искусственного продолжения материала.
Private Sub TestProviderStateSolutionExtension(ByRef stats As TMaterialTestStats)
    Dim provider As CMaterialModelProvider
    Set provider = TestProvider()

    Dim physicalConcrete As CMaterialDiagram
    Set physicalConcrete = provider.ConcreteMaterial(cpStrength)
    Dim stateConcrete As CMaterialDiagram
    Set stateConcrete = provider.ConcreteStateMaterial(cpStrength)
    AssertTrue stats, "state.extension.concrete.physicalClean", Not physicalConcrete.HasCompressionExtension
    AssertTrue stats, "state.extension.concrete.compression", stateConcrete.HasCompressionExtension
    AssertTrue stats, "state.extension.concrete.noTension", Not stateConcrete.HasTensionExtension
    AssertClose stats, "state.extension.concrete.physicalLimit", stateConcrete.PhysicalCompressionStrain, _
        physicalConcrete.PhysicalCompressionStrain, 0.000000000001

    Dim physicalSteel As CMaterialDiagram
    Set physicalSteel = provider.SteelMaterial(cpStrength)
    Dim stateSteel As CMaterialDiagram
    Set stateSteel = provider.SteelStateMaterial(cpStrength)
    AssertTrue stats, "state.extension.steel.compression", stateSteel.HasCompressionExtension
    AssertTrue stats, "state.extension.steel.tension", stateSteel.HasTensionExtension
    AssertClose stats, "state.extension.steel.physicalLimit", stateSteel.PhysicalTensionStrain, _
        physicalSteel.PhysicalTensionStrain, 0.000000000001
    AssertTrue stats, "state.extension.steel.plateauPhysical", Not stateSteel.IsInExtensionRange(0.01)
    AssertTrue stats, "state.extension.steel.beyondUltimate", stateSteel.IsInExtensionRange(0.03)

    Dim disabledProvider As CMaterialModelProvider
    Set disabledProvider = TestProvider("TwoLine", "Ignore", "TwoLine", "TwoLine", "TwoLine", False)
    AssertTrue stats, "state.extension.disabled", Not disabledProvider.SteelStateMaterial(cpStrength).HasTensionExtension
End Sub

' Проверяет требование ТЗ: пользовательские предельные деформации берутся из
' параметров Config, а не зашиты константами в построителе.
Private Sub TestUserStrainParametersAffectDiagrams(ByRef stats As TMaterialTestStats)
    Dim concrete As CConcreteMaterialParameters
    Set concrete = New CConcreteMaterialParameters
    concrete.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, _
        0.0017, 0.00009, 0.0022, 0.00011, 0.004, 0.0002, 14.6

    Dim steel As CSteelMaterialParameters
    Set steel = New CSteelMaterialParameters
    steel.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#, "Ribbed", _
        0.03, 0.031, 0.016, 0.017

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.InitializeFromParameters concrete, steel

    AssertClose stats, "material.userConcreteLimit", provider.ConcreteMaterial(cpStrength).PointStrain(1), -0.004, 0.000000000001
    AssertClose stats, "material.userConcreteTensionLimit", provider.ConcreteMaterial(cpMcrc).PointStrain(7), 0.0002, 0.000000000001
    AssertClose stats, "material.userSteelTwoLineCompressionLimit", provider.SteelMaterial(cpStrength).PointStrain(1), -0.03, 0.000000000001
    AssertClose stats, "material.userSteelTwoLineTensionLimit", provider.SteelMaterial(cpStrength).PointStrain(5), 0.031, 0.000000000001
End Sub

' Проверяет базовую защиту от некорректной диаграммы: точки должны иметь
' физически согласованные знаки и возрастающие деформации.
Private Sub TestInvalidParameters(ByRef stats As TMaterialTestStats)
    On Error GoTo ConcreteError
    Dim concrete As CMaterialDiagram
    Set concrete = New CMaterialDiagram
    concrete.Initialize 0.0015, -15.5, -0.0035, -15.5
    AssertTrue stats, "material.invalidConcrete", False
    GoTo SteelCheck

ConcreteError:
    On Error GoTo 0
    AssertTrue stats, "material.invalidConcrete", True
    Resume SteelCheck

SteelCheck:
    On Error GoTo SteelError
    Dim steel As CMaterialDiagram
    Set steel = New CMaterialDiagram
    steel.Initialize 0#, 350#, 0.025
    AssertTrue stats, "material.invalidSteel", False
    Exit Sub

SteelError:
    On Error GoTo 0
    AssertTrue stats, "material.invalidSteel", True
End Sub

Private Function TestProvider(Optional ByVal strengthConcreteDiagram As String = "TwoLine", _
        Optional ByVal strengthConcreteTension As String = "Ignore", _
        Optional ByVal strengthSteelDiagram As String = "TwoLine", _
        Optional ByVal mcrcSteelDiagram As String = "TwoLine", _
        Optional ByVal crackedSteelDiagram As String = "TwoLine", _
        Optional ByVal directStateDiagramExtension As Boolean = True) As CMaterialModelProvider
    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.InitializeFromParameters TestConcreteParameters(), TestSteelParameters(), _
        strengthConcreteDiagram, strengthConcreteTension, strengthSteelDiagram, _
        "ThreeLine", mcrcSteelDiagram, "TwoLine", crackedSteelDiagram, directStateDiagramExtension
    Set TestProvider = provider
End Function

Private Function TestConcreteParameters() As CConcreteMaterialParameters
    Dim parameters As CConcreteMaterialParameters
    Set parameters = New CConcreteMaterialParameters
    parameters.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, _
        0.0015, 0.00008, 0.002, 0.0001, 0.0035, 0.00015, 14.6
    Set TestConcreteParameters = parameters
End Function

Private Function TestSteelParameters() As CSteelMaterialParameters
    Dim parameters As CSteelMaterialParameters
    Set parameters = New CSteelMaterialParameters
    parameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#, "Ribbed", _
        0.025, 0.025, 0.015, 0.015
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
