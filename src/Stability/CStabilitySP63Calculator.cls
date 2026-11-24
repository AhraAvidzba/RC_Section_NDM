VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStabilitySP63Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Нормативный расчёт устойчивости по СП 63
' ==========================================================================
' Работает с двумя заранее подготовленными главными плоскостями.
' Геометрию и исходные нагрузки не читает из Excel; готовые моменты возвращает фасаду.

' Выбирает центр стержня, относительно которого СП 63 требует считать M1 и Ml1:
' наиболее растянутый или наименее сжатый стержень для текущего направления
' эксцентриситета. При отсутствии арматуры используется край бетонного контура,
' чтобы расчет не падал, но такая модель требует отдельной инженерной проверки.
Private Function RebarProjectionForMoment(ByVal section As CSectionModel, ByRef plane As TStabilityPlane) As Double
    If section Is Nothing Then
        RebarProjectionForMoment = StabilityConcreteFaceProjectionForMoment(plane)
        Exit Function
    End If
    If section.RebarCount <= 0 Then
        RebarProjectionForMoment = StabilityConcreteFaceProjectionForMoment(plane)
        Exit Function
    End If

    Dim chooseNegativeSide As Boolean
    chooseNegativeSide = (plane.Eccentricity >= 0#)

    Dim i As Long
    Dim projectionValue As Double
    For i = 1 To section.RebarCount
        projectionValue = plane.NormalX * (section.RebarX(i) - plane.CentroidX) + _
            plane.NormalY * (section.RebarY(i) - plane.CentroidY)
        If i = 1 Then
            RebarProjectionForMoment = projectionValue
        ElseIf chooseNegativeSide Then
            If projectionValue < RebarProjectionForMoment Then RebarProjectionForMoment = projectionValue
        Else
            If projectionValue > RebarProjectionForMoment Then RebarProjectionForMoment = projectionValue
        End If
    Next i
End Function

' Коэффициент длительности СП 63 считается не по моменту в центре сечения, а по
' моментам M1 и Ml1 относительно центра выбранного стержня арматуры.
Private Function PhiLSP63(ByVal phiLMode As String, ByVal section As CSectionModel, _
        ByVal nCompression As Double, ByVal sustainedCompression As Double, _
        ByRef plane As TStabilityPlane) As Double
    Dim rebarProjection As Double
    rebarProjection = RebarProjectionForMoment(section, plane)
    PhiLSP63 = StabilityPhiLFromMomentRatio(phiLMode, _
        StabilityMomentAboutProjection(plane.Moment, nCompression, rebarProjection), _
        StabilityMomentAboutProjection(plane.SustainedMoment, sustainedCompression, rebarProjection))
End Function

' ----------------------------- СП 63 --------------------------------------

' Расчет по СП 63: eta = 1/(1-N/Ncr), Ncr = pi^2*D/l0^2,
' D = kb*Eb*I + ks*Es*Is. Ветвь работает для произвольного дискретного
' сечения, потому что I, Is и h уже получены из CSectionModel.
Friend Sub CalculateSP63(ByVal section As CSectionModel, ByVal eb As Double, ByVal es As Double, _
        ByVal nCompression As Double, _
        ByVal sustainedCompression As Double, ByVal systemType As String, _
        ByVal phiLMode As String, ByVal ksValue As Double, ByVal deltaEMin As Double, _
        ByVal deltaEMax As Double, ByRef plane1 As TStabilityPlane, ByRef plane2 As TStabilityPlane)
    CalculateSP63Plane section, nCompression, sustainedCompression, eb, es, ksValue, systemType, phiLMode, deltaEMin, deltaEMax, plane1, "1"
    CalculateSP63Plane section, nCompression, sustainedCompression, eb, es, ksValue, systemType, phiLMode, deltaEMin, deltaEMax, plane2, "2"
End Sub

' Считает одну главную плоскость по формулам СП 63.
Private Sub CalculateSP63Plane(ByVal section As CSectionModel, ByVal nCompression As Double, _
        ByVal sustainedCompression As Double, ByVal eb As Double, ByVal es As Double, _
        ByVal ksValue As Double, ByVal systemType As String, ByVal phiLMode As String, ByVal deltaEMin As Double, _
        ByVal deltaEMax As Double, ByRef plane As TStabilityPlane, ByVal planeLabel As String)
    StabilityEnsurePositivePlaneGeometry plane, STABILITY_CODE_SP63, planeLabel

    plane.Eccentricity = StabilityInitialEccentricity(plane.Moment, nCompression, _
        plane.AccidentalEccentricity, systemType)
    plane.PhiL = PhiLSP63(phiLMode, section, nCompression, sustainedCompression, plane)

    Dim deltaE As Double
    deltaE = Abs(plane.Eccentricity) / plane.Depth
    If deltaEMin > 0# Then deltaE = GeomMax(deltaE, deltaEMin)
    If deltaEMax > 0# Then deltaE = GeomMin(deltaE, deltaEMax)

    Dim kb As Double
    kb = 0.15 / (plane.PhiL * (0.3 + deltaE))
    plane.DeltaValue = deltaE
    plane.Kb = kb
    plane.Ks = ksValue
    plane.StiffnessD = kb * eb * plane.InertiaConcrete + ksValue * es * plane.InertiaSteel
    plane.NcrForce = GEOM_PI * GEOM_PI * plane.StiffnessD / (plane.EffectiveLength * plane.EffectiveLength)
    plane.CriticalForce = plane.NcrForce
    plane.ReserveFactor = StabilityReserveByCriticalForce(plane.NcrForce, nCompression)
    If plane.Slenderness <= 14# Then
        plane.Eta = 1#
    Else
        plane.Eta = StabilityEtaByCriticalForce(plane.NcrForce, nCompression)
    End If
    plane.DesignMoment = StabilityDesignMomentWithEta(nCompression, plane.Eccentricity, plane.Eta)
    If plane.Slenderness <= 14# Then
        plane.CheckPassed = True
    Else
        plane.CheckPassed = (plane.ReserveFactor > 1# And plane.Eta > 0#)
    End If
    plane.BranchName = STABILITY_CODE_SP63
End Sub
