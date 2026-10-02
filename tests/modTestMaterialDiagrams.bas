Attribute VB_Name = "modTestMaterialDiagrams"
Option Explicit

' ==========================================================================
' Тесты диаграмм материалов
' ==========================================================================
' Проверяется целевая упрощенная архитектура:
' - CMaterialDiagram хранит готовые точки и выполняет только интерполяцию;
' - CMaterialModelProvider выбирает расчетный режим, I/II ГПС, TwoLine/ThreeLine
'   и строит точки диаграмм из параметров бетона и арматуры.
' Интерполяция проверяется отдельно от нормативного построения точек provider-ом.

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
    TestAudit02PhysicalDiagramPairs stats
    TestAudit02ExtensionBeyondTechnicalDefault stats
    TestAudit02ExtensionOverflowIsExplicit stats
    TestAudit03DiagramArrayInputs stats
    TestAudit03MaterialConfigBehavior stats

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

' Проверяет, что numerical extension появляется только у equilibrium-диаграмм.
' Физические Strength/Mcrc/CrackedNDS-диаграммы остаются без технических
' точек, поэтому физические пределы capacity и crack не расширяются.
Private Sub TestProviderStateSolutionExtension(ByRef stats As TMaterialTestStats)
    Dim provider As CMaterialModelProvider
    Set provider = TestProvider()

    Dim physicalConcrete As CMaterialDiagram
    Set physicalConcrete = provider.ConcreteMaterial(cpStrength)
    Dim stateConcrete As CMaterialDiagram
    Set stateConcrete = provider.ConcreteMaterialForEquilibrium(cpStrength)
    AssertTrue stats, "state.extension.concrete.physicalClean", Not physicalConcrete.HasCompressionExtension
    AssertTrue stats, "state.extension.concrete.compression", stateConcrete.HasCompressionExtension
    AssertTrue stats, "state.extension.concrete.noTension", Not stateConcrete.HasTensionExtension
    AssertClose stats, "state.extension.concrete.physicalLimit", stateConcrete.PhysicalCompressionStrain, _
        physicalConcrete.PhysicalCompressionStrain, 0.000000000001

    Dim physicalSteel As CMaterialDiagram
    Set physicalSteel = provider.SteelMaterial(cpStrength)
    Dim stateSteel As CMaterialDiagram
    Set stateSteel = provider.SteelMaterialForEquilibrium(cpStrength)
    AssertTrue stats, "state.extension.steel.compression", stateSteel.HasCompressionExtension
    AssertTrue stats, "state.extension.steel.tension", stateSteel.HasTensionExtension
    AssertClose stats, "state.extension.steel.physicalLimit", stateSteel.PhysicalTensionStrain, _
        physicalSteel.PhysicalTensionStrain, 0.000000000001
    AssertTrue stats, "state.extension.steel.plateauPhysical", Not stateSteel.IsInExtensionRange(0.01)
    AssertTrue stats, "state.extension.steel.beyondUltimate", stateSteel.IsInExtensionRange(0.03)

    Dim disabledProvider As CMaterialModelProvider
    Set disabledProvider = TestProvider("TwoLine", "Ignore", "TwoLine", "TwoLine", "TwoLine", False)
    AssertTrue stats, "state.extension.disabled", Not disabledProvider.SteelMaterialForEquilibrium(cpStrength).HasTensionExtension
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
    steel.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#, _
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

' Создает материалный provider с явно выбранными диаграммами ролей и Extension.
' Это позволяет проверить передачу селекторов независимо от Config книги.
Private Function TestProvider(Optional ByVal strengthConcreteDiagram As String = "TwoLine", _
        Optional ByVal strengthConcreteTension As String = "Ignore", _
        Optional ByVal strengthSteelDiagram As String = "TwoLine", _
        Optional ByVal mcrcSteelDiagram As String = "TwoLine", _
        Optional ByVal crackedSteelDiagram As String = "TwoLine", _
        Optional ByVal diagramExtensionEnabled As Boolean = True) As CMaterialModelProvider
    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.InitializeFromParameters TestConcreteParameters(), TestSteelParameters(), _
        strengthConcreteDiagram, strengthConcreteTension, strengthSteelDiagram, _
        "ThreeLine", mcrcSteelDiagram, "TwoLine", crackedSteelDiagram, diagramExtensionEnabled
    Set TestProvider = provider
End Function

' Возвращает фиксированные ULS/SLS сопротивления, модули и деформации бетона,
' используемые численными эталонами этого набора, а не нормативную марку.
Private Function TestConcreteParameters() As CConcreteMaterialParameters
    Dim parameters As CConcreteMaterialParameters
    Set parameters = New CConcreteMaterialParameters
    parameters.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, _
        0.0015, 0.00008, 0.002, 0.0001, 0.0035, 0.00015, 14.6
    Set TestConcreteParameters = parameters
End Function

' Возвращает фиксированные характеристики арматуры с разными предельными
' деформациями TwoLine/ThreeLine для проверки фактического выбора диаграммы.
Private Function TestSteelParameters() As CSteelMaterialParameters
    Dim parameters As CSteelMaterialParameters
    Set parameters = New CSteelMaterialParameters
    parameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#, _
        0.025, 0.025, 0.015, 0.015
    Set TestSteelParameters = parameters
End Function

' Однотипные assertions сохраняют конкретный test-ID и счетчики результатов.
' Численные проверки используют переданный абсолютный допуск без его ослабления.
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

' Добавляет диагностическую строку в итоговый отчет материалных тестов.
Private Sub AppendLine(ByRef stats As TMaterialTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' ==========================================================================
' ДЛЯ ТЕСТОВ: физическая модель и активные ветви общего Extension
' ==========================================================================

' Сравнивает узлы и внутренние точки всех сочетаний TwoLine/ThreeLine,
' ULS/SLS и Ignore/UseDiagram. Проверяет stress, tangent и совместный evaluator,
' а также фактическое продолжение только работающих ветвей конкретного spec.
Private Sub TestAudit02PhysicalDiagramPairs(ByRef stats As TMaterialTestStats)
    Dim diagramKind As Variant
    Dim valueSet As Variant
    Dim tensionKind As Variant
    Dim spec As CMaterialModelSpec
    Dim providerOn As CMaterialModelProvider
    Dim providerOff As CMaterialModelProvider
    Dim physical As CMaterialDiagram
    Dim extended As CMaterialDiagram
    Dim testKey As String
    For Each diagramKind In Array("TwoLine", "ThreeLine")
        For Each valueSet In Array("ULS(I)", "SLS(II)")
            For Each tensionKind In Array("Ignore", "UseDiagram")
                Set spec = New CMaterialModelSpec
                spec.Initialize CStr(valueSet), CStr(diagramKind), CStr(tensionKind), CStr(diagramKind)
                Set providerOn = TestProvider(diagramExtensionEnabled:=True)
                Set providerOff = TestProvider(diagramExtensionEnabled:=False)
                testKey = "audit02.material." & spec.SpecKey

                Set physical = providerOff.ConcreteMaterialForEquilibriumFromSpec(spec)
                Set extended = providerOn.ConcreteMaterialForEquilibriumFromSpec(spec)
                AssertPhysicalDiagramPair stats, testKey & ".concrete", physical, extended
                AssertTrue stats, testKey & ".concrete.compressionExtension", extended.HasCompressionExtension
                AssertClose stats, testKey & ".concrete.compressionSlope", _
                    extended.GetTangentModulus(physical.PhysicalCompressionStrain - 0.0001), 325#, 0.000000000001
                If CStr(tensionKind) = "Ignore" Then
                    AssertTrue stats, testKey & ".ignore.noExtension", Not extended.HasTensionExtension
                    AssertClose stats, testKey & ".ignore.stress", extended.GetStress(100#), 0#, 0.000000000001
                    AssertClose stats, testKey & ".ignore.tangent", extended.GetTangentModulus(100#), 0#, 0.000000000001
                    AssertTrue stats, testKey & ".ignore.range", extended.IsInPhysicalRange(100#)
                    AssertTrue stats, testKey & ".ignore.notUsed", Not extended.IsInExtensionRange(100#)
                Else
                    AssertTrue stats, testKey & ".tension.extension", extended.HasTensionExtension
                    AssertClose stats, testKey & ".tension.slope", _
                        extended.GetTangentModulus(physical.PhysicalTensionStrain + 0.0001), 325#, 0.000000000001
                    AssertTrue stats, testKey & ".tension.used", _
                        extended.IsInExtensionRange(physical.PhysicalTensionStrain + 0.0001)
                End If

                Set physical = providerOff.SteelMaterialForEquilibriumFromSpec(spec)
                Set extended = providerOn.SteelMaterialForEquilibriumFromSpec(spec)
                AssertPhysicalDiagramPair stats, testKey & ".steel", physical, extended
                AssertClose stats, testKey & ".steel.compressionSlope", _
                    extended.GetTangentModulus(physical.PhysicalCompressionStrain - 0.0001), 2000#, 0.000000000001
                AssertClose stats, testKey & ".steel.tensionSlope", _
                    extended.GetTangentModulus(physical.PhysicalTensionStrain + 0.0001), 2000#, 0.000000000001
            Next tensionKind
        Next valueSet
    Next diagramKind
End Sub

' Проверяет неизменность физического диапазона, узлов и середины каждого
' отрезка. Общий evaluator должен возвращать те же stress/tangent, что API.
Private Sub AssertPhysicalDiagramPair(ByRef stats As TMaterialTestStats, ByVal testKey As String, _
        ByVal physical As CMaterialDiagram, ByVal extended As CMaterialDiagram)
    AssertClose stats, testKey & ".compressionLimit", extended.PhysicalCompressionStrain, _
        physical.PhysicalCompressionStrain, 0.000000000001
    AssertClose stats, testKey & ".tensionLimit", extended.PhysicalTensionStrain, _
        physical.PhysicalTensionStrain, 0.000000000001
    Dim i As Long
    Dim strain As Double
    Dim stress As Double
    Dim tangent As Double
    For i = 1 To 2 * physical.PointCount - 1
        If i Mod 2 = 1 Then
            strain = physical.PointStrain((i + 1) \ 2)
        Else
            strain = 0.5 * (physical.PointStrain(i \ 2) + physical.PointStrain(i \ 2 + 1))
        End If
        AssertClose stats, testKey & ".stress." & CStr(i), extended.GetStress(strain), _
            physical.GetStress(strain), 0.000000000001
        AssertClose stats, testKey & ".tangent." & CStr(i), extended.GetTangentModulus(strain), _
            physical.GetTangentModulus(strain), 0.000000000001
        extended.EvaluateAtStrain strain, stress, tangent
        AssertClose stats, testKey & ".evaluateStress." & CStr(i), stress, physical.GetStress(strain), 0.000000000001
        AssertClose stats, testKey & ".evaluateTangent." & CStr(i), tangent, physical.GetTangentModulus(strain), 0.000000000001
        AssertTrue stats, testKey & ".physical." & CStr(i), extended.IsInPhysicalRange(strain)
        AssertTrue stats, testKey & ".notUsed." & CStr(i), Not extended.IsInExtensionRange(strain)
    Next i

    ' Внутренняя окрестность предела остается физической; малое округление
    ' за границей учитывается тем же допуском, а не включает extension-флаг.
    Dim boundary As Double, sign As Double
    Dim side As Variant
    For Each side In Array("compression", "tension")
        If CStr(side) = "compression" Then
            boundary = physical.PhysicalCompressionStrain
            sign = -1#
        Else
            If Not physical.HasPhysicalTensionLimit Then Exit For
            boundary = physical.PhysicalTensionStrain
            sign = 1#
        End If
        strain = boundary - sign * 0.000000001
        AssertClose stats, testKey & ".near." & CStr(side) & ".stress", _
            extended.GetStress(strain), physical.GetStress(strain), 0.000000000001
        AssertClose stats, testKey & ".near." & CStr(side) & ".tangent", _
            extended.GetTangentModulus(strain), physical.GetTangentModulus(strain), 0.000000000001
        strain = boundary + sign * 0.0000000000005
        AssertTrue stats, testKey & ".near." & CStr(side) & ".roundingPhysical", extended.IsInPhysicalRange(strain)
        AssertTrue stats, testKey & ".near." & CStr(side) & ".roundingNotUsed", Not extended.IsInExtensionRange(strain)
        strain = boundary + sign * 0.000000000002
        AssertTrue stats, testKey & ".near." & CStr(side) & ".outside", Not extended.IsInPhysicalRange(strain)
        AssertTrue stats, testKey & ".near." & CStr(side) & ".extensionUsed", extended.IsInExtensionRange(strain)
    Next side
End Sub

' Пользовательский физический предел за +/-10 не обрезается: технические
' узлы лежат дальше него, физическое плато и пределы остаются исходными.
Private Sub TestAudit02ExtensionBeyondTechnicalDefault(ByRef stats As TMaterialTestStats)
    Dim steel As CSteelMaterialParameters
    Set steel = New CSteelMaterialParameters
    steel.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#, 12#, 14#, 12#, 14#
    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.InitializeFromParameters TestConcreteParameters(), steel
    Dim extended As CMaterialDiagram
    Set extended = provider.SteelMaterialForEquilibrium(cpStrength)
    AssertTrue stats, "audit02.technical.compressionExtended", extended.HasCompressionExtension
    AssertTrue stats, "audit02.technical.tensionExtended", extended.HasTensionExtension
    AssertClose stats, "audit02.technical.left", extended.PointStrain(1), -24#, 0.000000000001
    AssertClose stats, "audit02.technical.right", extended.PointStrain(extended.PointCount), 28#, 0.000000000001
    AssertClose stats, "audit02.technical.physicalCompression", extended.PhysicalCompressionStrain, -12#, 0.000000000001
    AssertClose stats, "audit02.technical.physicalTension", extended.PhysicalTensionStrain, 14#, 0.000000000001
    AssertClose stats, "audit02.technical.physicalStress", extended.GetStress(13#), 350#, 0.000000000001
    AssertClose stats, "audit02.technical.extendedStress", extended.GetStress(15#), 2350#, 0.000000000001
End Sub

' Переполнение при построении технической ветви возвращается как явная
' ошибка построения, а не как урезанная диаграмма с ложным пределом.
Private Sub TestAudit02ExtensionOverflowIsExplicit(ByRef stats As TMaterialTestStats)
    On Error GoTo ExpectedError
    Dim steel As CSteelMaterialParameters
    Set steel = New CSteelMaterialParameters
    steel.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#, 1E+308, 1E+308, 1E+308, 1E+308
    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.InitializeFromParameters TestConcreteParameters(), steel
    Dim extended As CMaterialDiagram
    Set extended = provider.SteelMaterialForEquilibrium(cpStrength)
    AssertTrue stats, "audit02.technical.overflowRejected", False
    Exit Sub
ExpectedError:
    AssertTrue stats, "audit02.technical.overflowRejected", Err.Number = vbObjectError + 3245
End Sub

' ==========================================================================
' ДЛЯ ТЕСТОВ: контракты входных массивов готовой диаграммы
' ==========================================================================
' Выполняет тот же направленный контрпример отдельно от полной material suite.
' Ошибка инициализации не должна оставлять доступной частично замененную диаграмму.
Public Function RunAudit03DiagramArrayInputTests() As String
    Dim stats As TMaterialTestStats
    TestAudit03DiagramArrayInputs stats
    AppendLine stats, "TOTAL_AUDIT03_DIAGRAM_ARRAY_INPUT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03DiagramArrayInputTests = stats.Report
End Function

' Проверяет выделенность, наличие индексов 1..pointCount, повторные strain
' и цикл valid-invalid-valid. Это контракт API, не новая диаграмма/нормативный предел.
Private Sub TestAudit03DiagramArrayInputs(ByRef stats As TMaterialTestStats)
    Dim scenario As Variant, strains() As Double, stresses() As Double, diagram As CMaterialDiagram
    Dim pointCount As Long, expectedCode As Long, actualCode As Long, description As String
    Dim prefix As String
    On Error GoTo Failed
    For Each scenario In Array("CountTooSmall", "MissingStrain", "MissingStress", _
            "ShortStrain", "ShortStress", "BoundsStrain", "BoundsStress", "CountBeyondArrays", "Duplicate")
        Audit03FillDiagramArrays strains, stresses
        Set diagram = New CMaterialDiagram
        diagram.InitializeFromArrays strains, stresses, 3
        pointCount = 3
        expectedCode = vbObjectError + 3113
        Select Case CStr(scenario)
            Case "CountTooSmall": pointCount = 1: expectedCode = vbObjectError + 3100
            Case "MissingStrain": Erase strains
            Case "MissingStress": Erase stresses
            Case "ShortStrain": ReDim strains(1 To 2)
            Case "ShortStress": ReDim stresses(1 To 2)
            Case "BoundsStrain": ReDim strains(2 To 4)
            Case "BoundsStress": ReDim stresses(2 To 4)
            Case "CountBeyondArrays": pointCount = 4
            Case "Duplicate": strains(1) = strains(2): expectedCode = vbObjectError + 3111
        End Select
        prefix = "audit03.diagramArray." & CStr(scenario)
        actualCode = Audit03CaptureDiagramInitialize(diagram, strains, stresses, pointCount, description)
        AssertTrue stats, prefix & ".inputError", actualCode = expectedCode
        AssertTrue stats, prefix & ".reason", Len(description) > 0
        AssertTrue stats, prefix & ".noPoints", diagram.PointCount = 0
        AssertTrue stats, prefix & ".notUsable", Audit03CaptureDiagramStress(diagram) = vbObjectError + 3112
        AppendLine stats, "DIAGRAM_ARRAY_INPUT: case=" & CStr(scenario) & "|error=" & CStr(actualCode) & "|reason=" & description
        Audit03FillDiagramArrays strains, stresses
        diagram.InitializeFromArrays strains, stresses, 3
        AssertClose stats, prefix & ".restoredStress", diagram.GetStress(-0.00075), -7.75, 0.000000000001
        AssertTrue stats, prefix & ".restoredPoints", diagram.PointCount = 3
    Next scenario
    ' Дополнительный индекс 0 допустим, если нужные индексы 1..pointCount есть.
    ' Он не участвует в диаграмме; не меняем существующий контракт этого входа.
    Audit03FillDiagramArrays strains, stresses, 0
    strains(0) = -1#: stresses(0) = -999#
    diagram.InitializeFromArrays strains, stresses, 3
    AssertClose stats, "audit03.diagramArray.extraZeroIndex", diagram.GetStress(-0.00075), -7.75, 0.000000000001
    Exit Sub
Failed:
    AssertTrue stats, "audit03.diagramArray.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' Создает точки с нужными индексами для одинакового сравнения до/после guards.
' Порядок точек и численные эталоны не зависят от проверяемого actual-кода.
Private Sub Audit03FillDiagramArrays(ByRef strains() As Double, ByRef stresses() As Double, _
        Optional ByVal lowerIndex As Long = 1)
    ReDim strains(lowerIndex To 3)
    ReDim stresses(lowerIndex To 3)
    strains(1) = -0.0035: strains(2) = -0.0015: strains(3) = 0#
    stresses(1) = -15.5: stresses(2) = -15.5: stresses(3) = 0#
End Sub

' Перехватывает только ошибку InitializeFromArrays, чтобы не прятать отказ
' последующей проверки и не переносить Err от предыдущего сценария.
Private Function Audit03CaptureDiagramInitialize(ByVal diagram As CMaterialDiagram, _
        ByRef strains() As Double, ByRef stresses() As Double, ByVal pointCount As Long, _
        ByRef description As String) As Long
    On Error GoTo ExpectedError
    description = vbNullString
    diagram.InitializeFromArrays strains, stresses, pointCount
    Exit Function
ExpectedError:
    Audit03CaptureDiagramInitialize = Err.Number
    description = Err.Description
End Function

' Проверяет закрытие расчетного API после неуспешной инициализации.
' Успех или случайная ошибка индекса вместо EnsureInitialized не принимаются.
Private Function Audit03CaptureDiagramStress(ByVal diagram As CMaterialDiagram) As Long
    Dim stress As Double
    On Error GoTo ExpectedError
    stress = diagram.GetStress(-0.00075)
    Exit Function
ExpectedError:
    Audit03CaptureDiagramStress = Err.Number
End Function

' ==========================================================================
' ДЛЯ ТЕСТОВ: передача всех редактируемых параметров материалов из Config
' ==========================================================================

' Запускает адресную проверку отдельно от общей suite. Читает настоящие таблицы
' книги; эталонные значения и ожидаемые эффекты не берутся из actual-диаграмм.
Public Function RunAudit03MaterialConfigBehaviorTests() As String
    Dim stats As TMaterialTestStats
    TestAudit03MaterialConfigBehavior stats
    AppendLine stats, "TOTAL_AUDIT03_MATERIAL_CONFIG: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03MaterialConfigBehaviorTests = stats.Report
End Function

' Меняет по одному из 23 входов и проверяет все 16 профильных спецификаций.
' Неактивные параметры не меняют физическую диаграмму; отдельное Rb,mc2
' проверяется в калькуляторе продольных трещин, а не в sigma-epsilon диаграмме.
Private Sub TestAudit03MaterialConfigBehavior(ByRef stats As TMaterialTestStats)
    Dim concreteRange As Object, steelRange As Object, unitRange As Object
    Dim savedConcrete As Variant, savedSteel As Variant, savedUnits As Variant
    Dim values As Variant, changedValues As Variant, rowKeys As Variant, columns As Variant
    Dim i As Long, cell As Object, provider As CMaterialModelProvider, baseline As CMaterialModelProvider
    Dim valueSet As Variant, concreteKind As Variant, steelKind As Variant, tension As Variant
    Dim spec As CMaterialModelSpec, diagram As CMaterialDiagram, original As CMaterialDiagram
    Dim otherDiagram As CMaterialDiagram, originalOther As CMaterialDiagram
    Dim prefix As String, actual As Double, expected As Double, affected As Boolean, metricScale As Double
    Dim invalid As Variant, errorCode As Long, description As String, unitRow As Long
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo Failed
    Set concreteRange = ThisWorkbook.Names("rngConcreteMaterialParameters").RefersToRange
    Set steelRange = ThisWorkbook.Names("rngSteelMaterialParameters").RefersToRange
    Set unitRange = ThisWorkbook.Names("rngUnitSettings").RefersToRange
    savedConcrete = concreteRange.Formula
    savedSteel = steelRange.Formula
    savedUnits = unitRange.Formula
    values = Array(15.5, 1.1, 22#, 1.8, 14.6, 32500#, 31000#, 0.0015, 0.00008, _
        0.002, 0.0001, 0.0035, 0.00015, 350#, 340#, 390#, 380#, 200000#, 190000#, _
        0.025, 0.026, 0.015, 0.016)
    rowKeys = Array("Concrete.R.ULS(I)", "Concrete.R.ULS(I)", "Concrete.R.SLS(II)", _
        "Concrete.R.SLS(II)", "Concrete.Rb.mc2", "Concrete.E", "Concrete.E", _
        "Concrete.TwoLine.Eb1Red", "Concrete.TwoLine.Eb1Red", "Concrete.ThreeLine.Eb0", _
        "Concrete.ThreeLine.Eb0", "Concrete.TwoThreeLine.Eb2", "Concrete.TwoThreeLine.Eb2", _
        "Steel.R.ULS(I)", "Steel.R.ULS(I)", "Steel.R.SLS(II)", "Steel.R.SLS(II)", _
        "Steel.E", "Steel.E", "Steel.TwoLine.Es2", "Steel.TwoLine.Es2", _
        "Steel.ThreeLine.Es2", "Steel.ThreeLine.Es2")
    columns = Array(2, 3, 2, 3, 2, 2, 3, 2, 3, 2, 3, 2, 3, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2)
    For unitRow = 2 To unitRange.Rows.Count
        If StrComp(CStr(unitRange.Cells(unitRow, 1).Value2), "Stress", vbTextCompare) = 0 Then Exit For
    Next unitRow
    If unitRow > unitRange.Rows.Count Then Err.Raise vbObjectError + 9901, , "В fixture нет строки Stress."
    unitRange.Cells(unitRow, 2).Value2 = "MPa"
    For i = 0 To UBound(values)
        Set cell = Audit03MaterialConfigCell(concreteRange, steelRange, CStr(rowKeys(i)), CLng(columns(i)))
        cell.Value2 = values(i)
    Next i
    Set baseline = Audit03MaterialConfigProvider()
    For i = 0 To UBound(values)
        Set cell = Audit03MaterialConfigCell(concreteRange, steelRange, CStr(rowKeys(i)), CLng(columns(i)))
        changedValues = values
        changedValues(i) = CDbl(values(i)) * 1.1
        cell.Value2 = changedValues(i)
        Set provider = Audit03MaterialConfigProvider()
        For Each valueSet In Array("ULS(I)", "SLS(II)")
            For Each concreteKind In Array("TwoLine", "ThreeLine")
                For Each tension In Array("Ignore", "UseDiagram")
                    For Each steelKind In Array("TwoLine", "ThreeLine")
                        Set spec = New CMaterialModelSpec
                        spec.Initialize CStr(valueSet), CStr(concreteKind), CStr(tension), CStr(steelKind)
                        prefix = "audit03.materialConfig." & CStr(i) & "." & spec.SpecKey
                        If i < 13 Then
                            Set diagram = provider.ConcreteMaterialFromSpec(spec)
                            Set original = baseline.ConcreteMaterialFromSpec(spec)
                            Set otherDiagram = provider.SteelMaterialFromSpec(spec)
                            Set originalOther = baseline.SteelMaterialFromSpec(spec)
                        Else
                            Set diagram = provider.SteelMaterialFromSpec(spec)
                            Set original = baseline.SteelMaterialFromSpec(spec)
                            Set otherDiagram = provider.ConcreteMaterialFromSpec(spec)
                            Set originalOther = baseline.ConcreteMaterialFromSpec(spec)
                        End If
                        Audit03MaterialConfigMetric provider, diagram, i, changedValues, _
                            CStr(valueSet), CStr(concreteKind), CStr(tension), CStr(steelKind), actual, expected, affected
                        metricScale = Abs(expected)
                        If metricScale < 1# Then metricScale = 1#
                        AssertClose stats, prefix & ".expectedEffect", actual, expected, 0.000000001 * metricScale
                        AssertTrue stats, prefix & ".otherMaterialUnchanged", _
                            Audit03MaterialDiagramSignature(otherDiagram) = Audit03MaterialDiagramSignature(originalOther)
                        If affected Then
                            AssertTrue stats, prefix & ".activeChanged", _
                                Audit03MaterialDiagramSignature(diagram) <> Audit03MaterialDiagramSignature(original)
                        Else
                            AssertTrue stats, prefix & ".inactiveUnchanged", _
                                Audit03MaterialDiagramSignature(diagram) = Audit03MaterialDiagramSignature(original)
                        End If
                        If i = 5 Then AssertClose stats, prefix & ".referenceModulus", _
                            provider.ConcreteElasticModulusFromSpec(spec), CDbl(changedValues(i)), 0.000000001
                    Next steelKind
                Next tension
            Next concreteKind
        Next valueSet
        AppendLine stats, "MATERIAL_CONFIG_EFFECT: key=" & CStr(rowKeys(i)) & "|column=" & CStr(columns(i)) & _
            "|address=" & cell.Address & "|before=" & FormatNumberInvariant(CDbl(values(i))) & _
            "|after=" & FormatNumberInvariant(CDbl(changedValues(i))) & "|specs=16"
        For Each invalid In Array(0#, -1#, "TODO", vbNullString, CVErr(xlErrValue), "1e309")
            cell.Value2 = invalid
            errorCode = Audit03MaterialConfigError(description)
            AssertTrue stats, "audit03.materialConfig." & CStr(i) & ".invalid." & CStr(invalid), _
                errorCode <> 0 And errorCode <> 6 And errorCode <> 13
            AssertTrue stats, "audit03.materialConfig." & CStr(i) & ".reason." & CStr(invalid), Len(description) > 0
            AppendLine stats, "MATERIAL_CONFIG_INVALID: key=" & CStr(rowKeys(i)) & "|column=" & CStr(columns(i)) & _
                "|address=" & cell.Address & "|value=" & CStr(invalid) & "|error=" & CStr(errorCode) & "|reason=" & description
        Next invalid
        cell.Value2 = values(i)
        Set provider = Audit03MaterialConfigProvider()
        Set spec = New CMaterialModelSpec
        spec.Initialize "ULS(I)", "ThreeLine", "UseDiagram", "ThreeLine"
        AssertTrue stats, "audit03.materialConfig." & CStr(i) & ".recoveryConcrete", _
            Audit03MaterialDiagramSignature(provider.ConcreteMaterialFromSpec(spec)) = _
            Audit03MaterialDiagramSignature(baseline.ConcreteMaterialFromSpec(spec))
        AssertTrue stats, "audit03.materialConfig." & CStr(i) & ".recoverySteel", _
            Audit03MaterialDiagramSignature(provider.SteelMaterialFromSpec(spec)) = _
            Audit03MaterialDiagramSignature(baseline.SteelMaterialFromSpec(spec))
    Next i
    GoTo Cleanup
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    AssertTrue stats, "audit03.materialConfig.runtime." & CStr(failureNumber) & "." & failureDescription, False
Cleanup:
    On Error Resume Next
    If Not IsEmpty(savedConcrete) Then concreteRange.Formula = savedConcrete
    If Not IsEmpty(savedSteel) Then steelRange.Formula = savedSteel
    If Not IsEmpty(savedUnits) Then unitRange.Formula = savedUnits
    On Error GoTo 0
End Sub

' Находит ввод по смысловой строке и стороне материала, а не по текущему адресу
' Config. Отсутствующая строка является ошибкой fixture, не runtime-default.
Private Function Audit03MaterialConfigCell(ByVal concreteRange As Object, ByVal steelRange As Object, _
        ByVal rowKey As String, ByVal column As Long) As Object
    Dim source As Object, row As Long
    If Left$(rowKey, 9) = "Concrete." Then Set source = concreteRange Else Set source = steelRange
    For row = 2 To source.Rows.Count
        If StrComp(CStr(source.Cells(row, 1).Value2), rowKey, vbTextCompare) = 0 Then
            Set Audit03MaterialConfigCell = source.Cells(row, column)
            Exit Function
        End If
    Next row
    Err.Raise vbObjectError + 9902, , "В fixture нет параметра " & rowKey & "."
End Function

' Выполняет обычный production-маршрут чтения всей книги и преобразования
' единиц. InitializeFromParameters намеренно не используется в этой проверке.
Private Function Audit03MaterialConfigProvider() As CMaterialModelProvider
    Dim settings As CSystemSettingsReader, units As CUnitSystem, provider As CMaterialModelProvider
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem
    units.LoadFromSettings settings
    Set provider = New CMaterialModelProvider
    provider.Initialize settings, units
    Set Audit03MaterialConfigProvider = provider
End Function

' Перехватывает только контролируемую попытку прочитать неверный материал.
' Последующий recovery использует новый provider, чтобы не скрывать ошибку.
Private Function Audit03MaterialConfigError(ByRef description As String) As Long
    On Error GoTo ExpectedError
    Dim provider As CMaterialModelProvider
    description = vbNullString
    Set provider = Audit03MaterialConfigProvider()
    Exit Function
ExpectedError:
    Audit03MaterialConfigError = Err.Number
    description = Err.Description
End Function

' Выбирает наблюдаемую величину, чувствительную к конкретному входу. Эталон
' выводится из заданных чисел fixture; флаг affected относится только к самой
' физической диаграмме, а не к справочному модулю или продольной проверке.
Private Sub Audit03MaterialConfigMetric(ByVal provider As CMaterialModelProvider, ByVal diagram As CMaterialDiagram, _
        ByVal index As Long, ByRef v As Variant, ByVal valueSet As String, ByVal concreteKind As String, _
        ByVal tension As String, ByVal steelKind As String, ByRef actual As Double, _
        ByRef expected As Double, ByRef affected As Boolean)
    Dim rb As Double, rbt As Double, rs As Double, rsc As Double, three As Boolean, useTension As Boolean
    Dim longitudinal As CLongitudinalCrackCalculator, result As CLongitudinalCrackResult
    rb = CDbl(v(0)): rbt = CDbl(v(1)): rs = CDbl(v(13)): rsc = CDbl(v(14))
    If valueSet = "SLS(II)" Then rb = CDbl(v(2)): rbt = CDbl(v(3)): rs = CDbl(v(15)): rsc = CDbl(v(16))
    three = (concreteKind = "ThreeLine")
    useTension = (tension = "UseDiagram")
    affected = False
    Select Case index
        Case 0, 2
            actual = diagram.GetStress(-0.003): expected = -rb
            affected = ((index = 0) = (valueSet = "ULS(I)"))
        Case 1, 3
            actual = diagram.GetStress(0.00014): expected = 0#
            If useTension Then expected = rbt
            affected = useTension And ((index = 1) = (valueSet = "ULS(I)"))
        Case 4
            Set longitudinal = New CLongitudinalCrackCalculator
            Set result = longitudinal.CalculateFromStress(15#, provider.ConcreteParameters.RbMc2)
            actual = result.Utilization: expected = 15# / CDbl(v(4))
            If result.Status <> "OK" Then Err.Raise vbObjectError + 9903, , "После увеличения Rb,mc2 проверка должна пройти."
        Case 5
            actual = diagram.GetTangentModulus(-0.0000000001)
            expected = rb / CDbl(v(7)): affected = three
            If three Then expected = CDbl(v(5))
        Case 6
            actual = diagram.GetTangentModulus(0.0000000001): expected = 0#
            If useTension Then
                expected = rbt / CDbl(v(8))
                If three Then expected = CDbl(v(6))
            End If
            affected = three And useTension
        Case 7, 9
            actual = diagram.PointStrain(2): expected = -CDbl(v(7))
            If three Then expected = -CDbl(v(9))
            affected = ((index = 9) = three)
        Case 8, 10
            actual = 0#: expected = 0#
            If useTension Then
                actual = diagram.PointStrain(diagram.PointCount - 1): expected = CDbl(v(8))
                If three Then expected = CDbl(v(10))
            End If
            affected = useTension And ((index = 10) = three)
        Case 11
            actual = diagram.PhysicalCompressionStrain: expected = -CDbl(v(11)): affected = True
        Case 12
            actual = diagram.PhysicalTensionStrain: expected = 0#: affected = useTension
            If useTension Then expected = CDbl(v(12))
        Case 13, 15
            actual = diagram.GetStress(0.01): expected = rs
            If steelKind = "ThreeLine" Then expected = 1.1 * rs
            affected = ((index = 13) = (valueSet = "ULS(I)"))
        Case 14, 16
            actual = diagram.GetStress(-0.01): expected = -rsc
            If steelKind = "ThreeLine" Then expected = -1.1 * rsc
            affected = ((index = 14) = (valueSet = "ULS(I)"))
        Case 17
            actual = diagram.GetTangentModulus(0.0000000001): expected = CDbl(v(17)): affected = True
        Case 18
            actual = diagram.GetTangentModulus(-0.0000000001): expected = CDbl(v(18)): affected = True
        Case 19, 21
            actual = diagram.PhysicalTensionStrain: expected = CDbl(v(19))
            If steelKind = "ThreeLine" Then expected = CDbl(v(21))
            affected = ((index = 21) = (steelKind = "ThreeLine"))
        Case 20, 22
            actual = diagram.PhysicalCompressionStrain: expected = -CDbl(v(20))
            If steelKind = "ThreeLine" Then expected = -CDbl(v(22))
            affected = ((index = 22) = (steelKind = "ThreeLine"))
    End Select
End Sub

' Точный снимок физических узлов для проверки неизменности неактивной ветви.
' Не служит численным эталоном активного эффекта, который проверяется отдельно.
Private Function Audit03MaterialDiagramSignature(ByVal diagram As CMaterialDiagram) As String
    Dim i As Long, signature As String
    For i = 1 To diagram.PointCount
        signature = signature & CStr(diagram.PointStrain(i)) & ":" & CStr(diagram.PointStress(i)) & "|"
    Next i
    Audit03MaterialDiagramSignature = signature
End Function
