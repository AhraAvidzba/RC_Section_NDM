VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCrackWidthSP63Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Раскрытие нормальных трещин по СП 63
' ==========================================================================
' По готовому НДС выполняет нормативную подготовку, формулы и выбор определяющего результата.
' Образование трещины и равновесие не пересчитывает; геометрия и диагностика общие.
' Состояние принадлежит вызывающему фасаду и не сохраняется в этом объекте.

' Выполняет отдельную ветку СП для истинного центрального растяжения.
' Здесь нет нейтральной линии, поэтому Abt принимается равной всей площади
' бетона, а As - всей продольной арматуре, которая растянута в текущем НДС.
Friend Sub CalculateCentralTension(ByRef work As TCrackWidthWork, ByVal currentSolver As CSectionStateResult, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, _
        ByVal loadStateForClassification As CSectionLoadState)
    work.CentralTensionBranch = True
    work.Phi3 = ResolvePhi3(work, loadStateForClassification)

    If Not IsTensionLoad(loadStateForClassification) Then
        CrackFinishNoCrack work, "Центральная ветка: продольная сила не является растягивающей."
        Exit Sub
    End If

    Dim selected() As Boolean
    ReDim selected(1 To section.RebarCount)
    SelectAllTensionRebars work, currentSolver, section, selected
    If work.TensionRebarCount = 0 Then
        CrackFinishNoCrack work, "В центральном растяжении не найдено растянутой арматуры."
        Exit Sub
    End If

    Set work.InteractionRegion = work.GeometryQuery.ConcreteDomain
    work.Abt = work.InteractionRegion.Area
    work.RebarArea = SelectedRebarArea(section, selected)
    CompleteCrackFormula work, currentSolver, section, materialProvider.SteelMaterialFromSpec(work.CrackedStateSpec), selected
    If work.WidthInternalStatus <> rsUnset And Not work.CrackFormed Then Exit Sub
    If Not FinalizeWithPsiMode(work, section, materialProvider, selected, _
            "Расчет выполнен для центрального растяжения; Abt принята по всей площади бетона.") Then Exit Sub
End Sub

' Выполняет общий расчет для N + Mx + My при существующей нейтральной линии.
' Все геометрические величины строятся в координате s вдоль нормали к
' нейтральной линии, направленной в сторону растяжения.
Friend Sub CalculateGeneralBending(ByRef work As TCrackWidthWork, ByVal currentSolver As CSectionStateResult, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, _
        ByVal loadStateForClassification As CSectionLoadState)
    work.Phi3 = ResolvePhi3(work, loadStateForClassification)

    Dim nx As Double
    Dim ny As Double
    Dim normK As Double
    normK = Sqr(currentSolver.KappaX * currentSolver.KappaX + currentSolver.KappaY * currentSolver.KappaY)
    If normK <= CRACK_NEAR_ZERO_STRAIN Then ' При равномерной деформации нейтральной линии нет. Если порог образования трещины не выделен, определяем состояние по текущему НДС, не деля на нулевую кривизну.
        If CrackStrainAtConcreteCenter(currentSolver, section) > CRACK_NEAR_ZERO_STRAIN Then ' При почти равномерной плоскости проверяем деформацию в центре бетона, а не в удаленном начале WCS.
            CalculateCentralTension work, currentSolver, section, materialProvider, loadStateForClassification
        Else
            CrackFinishNoCrack work, "В текущем равномерном поле деформаций нет растянутой зоны для расчета раскрытия трещины."
        End If
        Exit Sub
    End If
    nx = currentSolver.KappaY / normK ' Из epsilon = epsilon0 + kappaX*y + kappaY*x нормаль к нейтральной линии имеет компоненты (kappaY, kappaX).
    ny = currentSolver.KappaX / normK

    Dim st As Double ' Наибольшая проекция бетона на нормаль к нейтральной линии, со стороны растяжения, мм.
    Dim sc As Double ' Наименьшая проекция бетона на ту же нормаль, с противоположной стороны, мм.
    ProjectionBoundsByConcreteContour work, section, nx, ny, st, sc
    work.SectionDepthH = st - sc
    If work.SectionDepthH <= CRACK_NEAR_ZERO_STRAIN Then
        CrackSetWidthInvalidInput work, "InputError: не удалось определить габарит сечения по нормали к нейтральной линии."
        Exit Sub
    End If

    Dim neutralS As Double
    neutralS = -currentSolver.Epsilon0 / normK ' Координата нулевой деформации на этой нормали; может лежать вне бетона, мм.
    work.TensionDepth = st - CrackMaxDouble(neutralS, sc)
    If work.TensionDepth <= CRACK_NEAR_ZERO_STRAIN Then
        CrackFinishNoCrack work, "В текущем состоянии нет растянутой бетонной зоны."
        Exit Sub
    End If

    Dim sBarMax As Double
    Dim coverRebarIndex As Long
    If Not FindMostTensionRebarProjection(currentSolver, section, nx, ny, sBarMax, coverRebarIndex) Then
        CrackFinishNoCrack work, "В расчетной растянутой зоне не найдено растянутой арматуры."
        Exit Sub
    End If
    work.CoverA = CoverDistanceForRebar(work, section, coverRebarIndex, nx, ny, st, sBarMax)
    If work.CoverA < 0# Then work.CoverA = 0#

    Dim zoneLowerS As Double
    zoneLowerS = TensionZoneLowerBound(work, st, sc, neutralS)
    work.EffectiveZoneDepth = st - zoneLowerS
    If work.EffectiveZoneDepth <= CRACK_NEAR_ZERO_STRAIN Then
        CrackFinishNoCrack work, "Эффективная растянутая зона имеет нулевую глубину."
        Exit Sub
    End If

    work.Abt = ConcreteAreaInZone(work, nx, ny, zoneLowerS)
    If work.Abt <= 0# Then
        CrackFinishNoCrack work, "В расчетную растянутую зону не попали бетонные элементы."
        Exit Sub
    End If

    Dim selected() As Boolean
    ReDim selected(1 To section.RebarCount)
    SelectRebarsInZone work, currentSolver, section, nx, ny, zoneLowerS, selected
    If work.TensionRebarCount = 0 Then
        CrackFinishNoCrack work, "В расчетную растянутую зону не попали растянутые стержни."
        Exit Sub
    End If

    CompleteCrackFormula work, currentSolver, section, materialProvider.SteelMaterialFromSpec(work.CrackedStateSpec), selected
    If work.WidthInternalStatus <> rsUnset And Not work.CrackFormed Then Exit Sub
    If Not FinalizeWithPsiMode(work, section, materialProvider, selected, _
            "Расчет выполнен для произвольного сечения по общей НДМ-ветке.") Then Exit Sub
End Sub

' Завершает подготовку нормативной формулы раскрытия трещин после выбора
' растянутой зоны и арматуры. Метод собирает As, Abt, d_s,eq, sigma_s и
' расстояние между трещинами l_s; сам коэффициент psi_s уточняется ниже.
Private Sub CompleteCrackFormula(ByRef work As TCrackWidthWork, ByVal solver As CSectionStateResult, ByVal section As CSectionModel, _
        ByVal steelMaterial As Object, ByRef selected() As Boolean)
    work.Phi1 = work.Phi1Input
    work.Phi2 = work.Phi2Input
    If work.Phi3 <= 0# Then work.Phi3 = work.Phi3Input
    work.PsiS = work.PsiSInput

    work.RebarArea = SelectedRebarArea(section, selected)
    work.TensionRebarIds = SelectedRebarIds(section, selected)
    work.DsEquivalent = EquivalentDiameter(section, selected)
    work.SigmaS = AverageSelectedSteelStress(work, solver.Epsilon0, solver.KappaX, solver.KappaY, section, steelMaterial, selected)
    If work.RebarArea <= 0# Or work.Abt <= 0# Or work.DsEquivalent <= 0# Or work.SigmaS <= 0# Then
        CrackFinishNoCrack work, "Недостаточно растянутой арматуры или бетона для расчета раскрытия трещины."
        Exit Sub
    End If

    work.LsRaw = 0.5 * (work.Abt / work.RebarArea) * work.DsEquivalent ' Исходное расстояние между трещинами, мм, до применения нижней и верхней границ.
    work.Ls = LimitedCrackSpacing(work, work.LsRaw, work.DsEquivalent)
End Sub

' Завершает расчет раскрытия с учетом режима psi_s.
' Если пригодная точка образования отсутствует, любой режим использует
' psi_s = 1 с предупреждением; текущий НДС остается обязательным.
' При доступной точке User использует заданный пользователем psi_s.
' Если выбранная арматура в готовом Post не растянута, уточнять коэффициент
' по этому состоянию нельзя: при текущем растяжении все режимы используют
' psi_s=1 с предупреждением, но формула ширины продолжает выполняться.
' Auto сначала делает проверку с psi_s = 1 и уточняет коэффициент только при
' непрохождении. AlwaysCalc сразу пытается применить уточненный psi_s. Точка
' lambda_crc и состояния PreCrackState/PostCrackState берутся из предварительной
' проверки образования, если она их получила. Повторный поиск здесь не выполняется.
Private Function FinalizeWithPsiMode(ByRef work As TCrackWidthWork, ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByRef selected() As Boolean, _
        ByVal finalReason As String) As Boolean
    If Not work.FallbackPsi1 And Not work.Formation.FallbackPsi1 Then
        CaptureSigmaSCrcIfAvailable work, section, materialProvider, selected
        If Not work.PostHasSelectedTension Then
            work.FallbackPsi1 = True
            work.PsiFallbackReason = "Предупреждение: в состоянии образования трещины выбранная арматура не растянута (напряжения неположительные); " & _
                "psi_s принят равным 1, уменьшение раскрытия не применяется. Ширина рассчитана по растянутой арматуре текущего НДС."
        End If
    End If
    If work.FallbackPsi1 Then
        work.PsiS = 1#
        If Len(work.PsiFallbackReason) = 0 Then work.PsiFallbackReason = _
            "Предупреждение: пригодное состояние сразу после образования трещины недоступно; " & _
            "psi_s принят равным 1, уменьшение раскрытия не применяется. Расчет выполнен по НДС заданного сочетания."
        FinalizeWidth work, finalReason & " " & work.PsiFallbackReason
        FinalizeWithPsiMode = work.Converged
        Exit Function
    End If
    If StrComp(work.PsiMode, CRACK_PSI_MODE_USER, vbTextCompare) = 0 Then
        work.PsiS = work.PsiSInput
        CaptureSigmaSCrcIfAvailable work, section, materialProvider, selected
        FinalizeWidth work, finalReason & " Psi_s принят из настройки SLS.Crack.PsiS."
        FinalizeWithPsiMode = work.Converged
        Exit Function
    End If

    If StrComp(work.PsiMode, CRACK_PSI_MODE_ALWAYS_CALC, vbTextCompare) = 0 Then
        If ResolveAutoPsiIfPossible(work, section, materialProvider, selected, _
                "AlwaysCalc psi_s") Then
            FinalizeWidth work, finalReason & AutomaticPsiComment(work)
        ElseIf work.Formation.FallbackPsi1 Then
            work.PsiS = 1#
            FinalizeWidth work, finalReason & " Psi_s оставлен равным 1, потому что выбранный путь не дает отдельной точки sigma_s,crc."
        Else
            Exit Function
        End If
        FinalizeWithPsiMode = work.Converged
        Exit Function
    End If

    work.PsiS = 1#
    CaptureSigmaSCrcIfAvailable work, section, materialProvider, selected
    FinalizeWidth work, finalReason & " Auto psi_s: первая проверка выполнена с psi_s=1."
    If Not work.Converged Or Not work.CrackFormed Then Exit Function

    If work.AllowableCrackWidth > 0# And work.CrackWidth <= work.AllowableCrackWidth Then
        CrackAppendDiagnostic work, "Auto psi_s: первая проверка с psi_s=1 проходит; sigma_s,crc для уточнения psi_s не потребовался."
        FinalizeWithPsiMode = True
        Exit Function
    End If

    If work.Formation.FallbackPsi1 Then
        CrackAppendDiagnostic work, "Auto psi_s: выбранный путь образования трещины не дает отдельной точки sigma_s,crc; psi_s оставлен равным 1."
        FinalizeWithPsiMode = True
        Exit Function
    End If

    CrackAppendDiagnostic work, "Auto psi_s: первая проверка с psi_s=1 не проходит; используется ранее найденный PostCrackState."
    work.Converged = False

    If Not ResolveAutoPsiIfPossible(work, section, materialProvider, selected, _
            "Auto psi_s") Then Exit Function
    FinalizeWidth work, finalReason & AutomaticPsiComment(work)
    FinalizeWithPsiMode = work.Converged
End Function

' Объясняет фактически примененный коэффициент после подготовки PostCrackState.
' Неположительное среднее напряжение не является несходимостью и не уменьшает
' раскрытие; текст результата должен явно объяснять принятие psi_s = 1.
Private Function AutomaticPsiComment(ByRef work As TCrackWidthWork) As String
    If work.SigmaSCrc <= CRACK_NEAR_ZERO_STRESS Then
        AutomaticPsiComment = " Среднее напряжение sigma_s,crc по выбранным стержням неположительное или практически нулевое; " & _
            "psi_s принят равным 1, уменьшение раскрытия не применяется."
    Else
        AutomaticPsiComment = " Psi_s рассчитан по состоянию сразу после образования трещины."
    End If
End Function

' Готовит и применяет автоматическое значение psi_s по sigma_s,crc.
' Метод не запускает solve и не достраивает PostCrackState: на этом этапе
' используются только данные, уже подготовленные проверкой образования трещины.
' Если sigma_s,crc получить нельзя, psi_s остается равным 1 с диагностикой.
Private Function ResolveAutoPsiIfPossible(ByRef work As TCrackWidthWork, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByRef selected() As Boolean, _
        ByVal contextText As String) As Boolean
    If work.Formation.FallbackPsi1 Then
        CrackAppendDiagnostic work, contextText & ": выбранный путь не дает отдельной точки sigma_s,crc."
        Exit Function
    End If

    If work.Formation.LambdaCrc < 0# Then
        CrackSetWidthInvalidInput work, "InputError: lambda_crc предварительной проверки образования трещины не определена."
        Exit Function
    End If

    If work.SigmaSCrc <= CRACK_NEAR_ZERO_STRESS Then _
        CaptureSigmaSCrcIfAvailable work, section, materialProvider, selected

    If work.SigmaSCrc <= CRACK_NEAR_ZERO_STRESS Then
        CrackAppendDiagnostic work, contextText & ": среднее sigma_s,crc неположительное или практически нулевое; psi_s оставлен равным 1."
        work.PsiS = 1#
        ResolveAutoPsiIfPossible = True
        Exit Function
    End If

    work.PsiS = AutoPsi(work.SigmaS, work.SigmaSCrc)
    ResolveAutoPsiIfPossible = True
End Function

' Получает справочное sigma_s,crc по сохранённому НДС после образования трещины.
' Новое равновесие не решается и действующий объект решателя не требуется.
' Если отдельная пороговая точка не определена, соответствующего sigma_s,crc нет.
Private Sub CaptureSigmaSCrcIfAvailable(ByRef work As TCrackWidthWork, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByRef selected() As Boolean)
    If work.SigmaSCrc > CRACK_NEAR_ZERO_STRESS Then Exit Sub
    If work.Formation.FallbackPsi1 Then Exit Sub
    Dim postState As CSectionStateResult
    Set postState = work.Formation.PostCrackState
    If postState Is Nothing Then Exit Sub
    If Not postState.Converged Then Exit Sub
    If Not postState.WithinPhysicalRange Then Exit Sub

    Dim i As Long, steelMaterial As CMaterialDiagram, strain As Double
    Set steelMaterial = materialProvider.SteelMaterialFromSpec(work.CrackedStateSpec)
    For i = 1 To section.RebarCount
        If selected(i) Then
            strain = postState.Epsilon0 + postState.KappaX * section.RebarY(i) + postState.KappaY * section.RebarX(i)
            If steelMaterial.GetStress(strain) > CRACK_NEAR_ZERO_STRESS Then work.PostHasSelectedTension = True
        End If
    Next i

    work.SigmaSCrc = AverageSelectedSteelStress(work, postState.Epsilon0, postState.KappaX, postState.KappaY, section, _
        materialProvider.SteelMaterialFromSpec(work.CrackedStateSpec), selected, False, IncludeCompressedStressForSigmaSCrc(work))
End Sub

' Завершает расчет по формуле СП 63 п. 8.2.15 для продолжительного раскрытия.
Private Sub FinalizeWidth(ByRef work As TCrackWidthWork, ByVal reason As String)
    If work.SigmaS <= 0# Or work.Ls <= 0# Then
        CrackFinishNoCrack work, "Напряжение растянутой арматуры или ls равны нулю."
        Exit Sub
    End If

    work.FormulaCalculated = True
    If Len(work.SpacingLimitWarning) > 0 Then reason = reason & " " & work.SpacingLimitWarning
    work.CrackWidth = CrackWidthFromData(work.Phi1, work.Phi2, work.Phi3, work.PsiS, work.SigmaS, work.SteelEs, work.Ls)
    work.Utilization = CrackUtilizationFromData(work.CrackWidth, work.AllowableCrackWidth)
    work.CrackFormed = True
    work.Converged = True
    If work.AllowableCrackWidth > 0# And work.CrackWidth > work.AllowableCrackWidth Then
        CrackSetWidthResult work, rsCheckFailed, rcCrackWidthExceeded, _
            "Раскрытие нормальной трещины " & CrackFormatNumberInvariant(work.CrackWidth) & _
            " мм превышает допускаемое " & CrackFormatNumberInvariant(work.AllowableCrackWidth) & " мм. " & reason
    Else
        If work.FallbackPsi1 Or work.HasSpacingLimitWarning Then
            CrackSetWidthResult work, rsSuccessWithWarning, rcCheckPassed, reason
        Else
            CrackSetWidthResult work, rsSuccess, rcCheckPassed, reason
        End If
    End If
    CrackAppendDiagnostic work, "a_crc=" & CrackFormatNumberInvariant(work.CrackWidth) & _
        "; phi1=" & CrackFormatNumberInvariant(work.Phi1) & _
        "; phi2=" & CrackFormatNumberInvariant(work.Phi2) & _
        "; phi3=" & CrackFormatNumberInvariant(work.Phi3) & _
        "; psi_s=" & CrackFormatNumberInvariant(work.PsiS) & _
        "; sigma_s=" & CrackFormatNumberInvariant(work.SigmaS) & _
        "; ls=" & CrackFormatNumberInvariant(work.Ls)
End Sub

' Вычисляет a_crc только по переданным коэффициентам, напряжению,
' модулю и расстоянию между трещинами. Состав арматуры, НДС
' и статус проверки этим методом не определяются.
Friend Function CrackWidthFromData(ByVal phi1 As Double, ByVal phi2 As Double, _
        ByVal phi3 As Double, ByVal psiS As Double, ByVal sigmaS As Double, _
        ByVal steelEs As Double, ByVal crackSpacing As Double) As Double
    If steelEs <= 0# Then Exit Function
    CrackWidthFromData = phi1 * phi2 * phi3 * psiS * (sigmaS / steelEs) * crackSpacing
End Function

' Выбирает все растянутые стержни для центрального растяжения.
Private Sub SelectAllTensionRebars(ByRef work As TCrackWidthWork, ByVal solver As CSectionStateResult, ByVal section As CSectionModel, ByRef selected() As Boolean)
    Dim i As Long
    For i = 1 To section.RebarCount
        If RebarStrain(solver, section, i) > 0# Then
            selected(i) = True
            work.TensionRebarCount = work.TensionRebarCount + 1
        End If
    Next i
End Sub

' Находит крайнюю растянутую арматуру по нормали к нейтральной линии.
Private Function FindMostTensionRebarProjection(ByVal solver As CSectionStateResult, ByVal section As CSectionModel, _
        ByVal nx As Double, ByVal ny As Double, ByRef sBarMax As Double, ByRef rebarIndex As Long) As Boolean
    Dim i As Long
    For i = 1 To section.RebarCount
        If RebarStrain(solver, section, i) > 0# Then
            Dim s As Double
            s = Projection(section.RebarX(i), section.RebarY(i), nx, ny)
            If Not FindMostTensionRebarProjection Or s > sBarMax Then sBarMax = s
            If s >= sBarMax - CRACK_NEAR_ZERO_STRAIN Then rebarIndex = i
            FindMostTensionRebarProjection = True
        End If
    Next i
End Function

' Выбирает смысл расстояния a_s по настройке CoverMode. NearestContour
' использует локальный выход луча от стержня к бетонному контуру;
' GlobalExtreme считает расстояние до опорной линии всего сечения.
' Если локальный контур не определен, используется расстояние GlobalExtreme.
Private Function CoverDistanceForRebar(ByRef work As TCrackWidthWork, ByVal section As CSectionModel, ByVal rebarIndex As Long, _
        ByVal nx As Double, ByVal ny As Double, ByVal st As Double, ByVal sBarMax As Double) As Double
    If StrComp(work.CoverMode, CRACK_COVER_MODE_GLOBAL_EXTREME, vbTextCompare) = 0 Then
        CoverDistanceForRebar = st - sBarMax
        Exit Function
    End If

    CoverDistanceForRebar = LocalConcreteCoverForRebar(work, section, rebarIndex, nx, ny)
    If CoverDistanceForRebar < 0# Then CoverDistanceForRebar = st - sBarMax
End Function

' Для a_s нужна не глобальная крайняя проекция всего сечения, а расстояние от
' оси выбранного растянутого стержня до ближайшего бетонного контура по нормали
' к нейтральной линии. Это особенно важно для Г-сечений и импортированных
' повернутых сеток, где глобальная опорная линия может лежать на другом выступе.
Private Function LocalConcreteCoverForRebar(ByRef work As TCrackWidthWork, ByVal section As CSectionModel, ByVal rebarIndex As Long, _
        ByVal nx As Double, ByVal ny As Double) As Double
    If rebarIndex <= 0 Then
        LocalConcreteCoverForRebar = -1#
        Exit Function
    End If

    LocalConcreteCoverForRebar = work.GeometryQuery.ConcreteCoverFromPointAlongDirection( _
        section.RebarX(rebarIndex), section.RebarY(rebarIndex), nx, ny)
End Function

' Определяет нижнюю границу расчетной зоны в координате s.
' FullTension берет всю фактически растянутую часть до нейтральной линии.
' Effective задает нормативную глубину h_bt = min(max(x_t, 2a), 0.5h).
' Если фактическая растянутая зона x_t меньше 2a, полоса может геометрически
' зайти за нейтральную линию. Для Abt это намеренно учитывается: готовая область
' отсекается по полосе h_bt, а не по знаку деформации центра бетонного КЭ.
' Арматура As остается только растянутой, потому что sigma_s в формуле раскрытия
' относится к работающей растянутой арматуре между трещинами.
Private Function TensionZoneLowerBound(ByRef work As TCrackWidthWork, ByVal st As Double, ByVal sc As Double, ByVal neutralS As Double) As Double
    Dim lower As Double
    If StrComp(work.TensionZoneMode, "FullTension", vbTextCompare) = 0 Then
        lower = CrackMaxDouble(neutralS, sc)
    Else
        Dim effectiveDepth As Double
        effectiveDepth = CrackMinDouble(CrackMaxDouble(work.TensionDepth, 2# * work.CoverA), 0.5 * work.SectionDepthH)
        If effectiveDepth <= CRACK_NEAR_ZERO_STRAIN Then effectiveDepth = CrackMinDouble(work.TensionDepth, 0.5 * work.SectionDepthH)
        lower = st - effectiveDepth
        lower = CrackMaxDouble(lower, sc)
    End If
    TensionZoneLowerBound = lower
End Function

' Отсекает фактическую бетонную область по нормативной границе зоны СП 63.
' Площадь частичных элементов и отверстий принадлежит общему движку; эта же
' конечная область сохраняется для Results и экспорта, без второго построения.
Private Function ConcreteAreaInZone(ByRef work As TCrackWidthWork, ByVal nx As Double, ByVal ny As Double, ByVal lowerS As Double) As Double
    Set work.InteractionRegion = work.GeometryQuery.ClipHalfPlane(work.GeometryQuery.ConcreteDomain, nx, ny, lowerS)
    ConcreteAreaInZone = work.InteractionRegion.Area
End Function

' Формирует набор стержней As по той же полосе, которая принята для Abt,
' но с обязательным физическим фильтром по растяжению стержня. Поэтому при
' Effective и срабатывании ограничения 2a бетонная полоса может зайти за
' нейтральную линию, однако сжатые стержни в As, sigma_s и d_s,eq не попадают.
Private Sub SelectRebarsInZone(ByRef work As TCrackWidthWork, ByVal solver As CSectionStateResult, ByVal section As CSectionModel, _
        ByVal nx As Double, ByVal ny As Double, ByVal lowerS As Double, ByRef selected() As Boolean)
    Dim i As Long
    For i = 1 To section.RebarCount
        If RebarStrain(solver, section, i) > 0# Then
            If work.GeometryQuery.ContainsPoint(work.InteractionRegion, section.RebarX(i), section.RebarY(i)) Then
                selected(i) = True
                work.TensionRebarCount = work.TensionRebarCount + 1
            End If
        End If
    Next i
End Sub

' Получает крайние проекции из готовой бетонной области: по точному контуру,
' если он задан, иначе по границам бетонных элементов. Для устойчивости
' проекции сетки вычисляет отдельно CSectionPropertiesCalculator.
Private Sub ProjectionBoundsByConcreteContour(ByRef work As TCrackWidthWork, ByVal section As CSectionModel, _
        ByVal nx As Double, ByVal ny As Double, ByRef st As Double, ByRef sc As Double)
    work.GeometryQuery.ProjectionBounds work.GeometryQuery.ConcreteDomain, nx, ny, sc, st
End Sub

' Считает средневзвешенное напряжение выбранной арматуры. В формулу раскрытия
' идет одно глобальное sigma_s, а не локальные значения по отдельным зонам.
Private Function AverageSelectedSteelStress(ByRef work As TCrackWidthWork, ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double, _
        ByVal section As CSectionModel, _
        ByVal steelMaterial As Object, ByRef selected() As Boolean, _
        Optional ByVal trackCurrentMaxima As Boolean = True, _
        Optional ByVal includeCompressedStress As Boolean = False) As Double
    Dim weightedStress As Double
    Dim areaSum As Double
    Dim i As Long
    For i = 1 To section.RebarCount
        If selected(i) Then
            Dim strain As Double
            Dim stress As Double
            strain = epsilon0 + kappaX * section.RebarY(i) + kappaY * section.RebarX(i)
            stress = steelMaterial.GetStress(strain)
            If (Not includeCompressedStress) And stress < 0# Then stress = 0# ' Сжатый вклад обнуляется, но площадь выбранного стержня остаётся в знаменателе среднего напряжения.
            weightedStress = weightedStress + section.RebarArea(i) * stress
            areaSum = areaSum + section.RebarArea(i)
            If trackCurrentMaxima Then TrackMaxSteel work, section, i, strain, stress
        End If
    Next i
    If areaSum > 0# Then AverageSelectedSteelStress = weightedStress / areaSum
End Function

' Возвращает режим усреднения именно для sigma_s,crc. Текущее sigma_s основной
' формулы раскрытия считается только по растягивающей части
' выбранной группы; этот переключатель нужен для состояния PostCrackState,
' где часть той же группы может перейти в сжатие сразу после образования трещины.
Private Function IncludeCompressedStressForSigmaSCrc(ByRef work As TCrackWidthWork) As Boolean
    IncludeCompressedStressForSigmaSCrc = (StrComp(work.SigmaSCrcAveragingMode, _
        CRACK_SIGMA_S_CRC_AVERAGE_ALL_SELECTED, vbTextCompare) = 0)
End Function

' Считает площадь выбранной арматуры.
Private Function SelectedRebarArea(ByVal section As CSectionModel, ByRef selected() As Boolean) As Double
    Dim i As Long
    For i = 1 To section.RebarCount
        If selected(i) Then SelectedRebarArea = SelectedRebarArea + section.RebarArea(i)
    Next i
End Function

' Перечисляет имена выбранных стержней для контрольной таблицы напряжений.
Private Function SelectedRebarIds(ByVal section As CSectionModel, ByRef selected() As Boolean) As String
    Dim result As String
    Dim i As Long
    For i = 1 To section.RebarCount
        If selected(i) Then
            If Len(result) > 0 Then result = result & ", "
            result = result & section.RebarID(i)
        End If
    Next i
    SelectedRebarIds = result
End Function

' Для разных диаметров используется принятое обобщение эквивалентного диаметра:
' ds,eq = sum(n_i*d_i^2) / sum(n_i*d_i). В дискретной модели каждый стержень
' считается отдельным i, поэтому n_i фактически равен количеству совпадающих d.
Private Function EquivalentDiameter(ByVal section As CSectionModel, ByRef selected() As Boolean) As Double
    Dim numerator As Double
    Dim denominator As Double
    Dim i As Long
    For i = 1 To section.RebarCount
        If selected(i) Then
            Dim d As Double
            d = section.RebarDiameter(i)
            If d <= 0# Then d = Sqr(4# * section.RebarArea(i) / GEOM_PI)
            numerator = numerator + d * d
            denominator = denominator + d
        End If
    Next i
    If denominator > 0# Then EquivalentDiameter = numerator / denominator
End Function

' Ограничивает ls снизу, затем сверху по принятой расчетной последовательности.
' Если нижняя граница больше верхней, сохраняет верхнюю и инженерное
' предупреждение для результата формулы; это не ошибка поиска равновесия.
Private Function LimitedCrackSpacing(ByRef work As TCrackWidthWork, ByVal rawLs As Double, ByVal ds As Double) As Double
    Dim lower As Double
    Dim upper As Double
    lower = CrackMaxDouble(10# * ds, 100#)
    upper = CrackMinDouble(40# * ds, 400#)
    If lower > upper Then
        work.HasSpacingLimitWarning = True
        work.SpacingLimitWarning = "Предупреждение: ограничения расстояния между трещинами противоречат друг другу: " & _
            "нижняя граница " & CrackFormatNumberInvariant(lower) & " мм выше верхней " & _
            CrackFormatNumberInvariant(upper) & " мм при d_s,eq = " & CrackFormatNumberInvariant(ds) & _
            " мм. Принято ls = " & CrackFormatNumberInvariant(upper) & " мм по правилу ограничения сверху."
        CrackAppendDiagnostic work, work.SpacingLimitWarning
        LimitedCrackSpacing = upper
    Else
        LimitedCrackSpacing = CrackMinDouble(CrackMaxDouble(rawLs, lower), upper)
    End If
End Function

' Формула psi_s = 1 - 0.8*sigma_s,crc/sigma_s. Ограничение 0...1 является
' численной защитой от отрицательного psi_s при почти равных напряжениях или
' случайных малых несоответствиях solve-ов.
Friend Function AutoPsi(ByVal sigmaS As Double, ByVal sigmaSCrc As Double) As Double
    If sigmaS <= 0# Then
        AutoPsi = 1#
    Else
        AutoPsi = 1# - 0.8 * sigmaSCrc / sigmaS
        If AutoPsi < 0# Then AutoPsi = 0#
        If AutoPsi > 1# Then AutoPsi = 1#
    End If
End Function

' Суммарная площадь бетона модели.
Private Function TotalConcreteArea(ByVal section As CSectionModel) As Double
    Dim i As Long
    For i = 1 To section.ConcreteCount
        TotalConcreteArea = TotalConcreteArea + section.ConcreteArea(i)
    Next i
End Function

' Деформация бетонного КЭ.
Private Function ConcreteStrain(ByVal solver As CSectionStateResult, ByVal section As CSectionModel, ByVal index As Long) As Double
    ConcreteStrain = solver.Epsilon0 + solver.KappaX * section.ConcreteY(index) + solver.KappaY * section.ConcreteX(index)
End Function

' Деформация стержня арматуры.
Private Function RebarStrain(ByVal solver As CSectionStateResult, ByVal section As CSectionModel, ByVal index As Long) As Double
    RebarStrain = solver.Epsilon0 + solver.KappaX * section.RebarY(index) + solver.KappaY * section.RebarX(index)
End Function

' Проекция точки на нормаль к нейтральной линии.
Private Function Projection(ByVal xCoord As Double, ByVal yCoord As Double, ByVal nx As Double, ByVal ny As Double) As Double
    Projection = nx * xCoord + ny * yCoord
End Function

' Обновляет максимум деформации и напряжения арматуры для диагностики.
Private Sub TrackMaxSteel(ByRef work As TCrackWidthWork, ByVal section As CSectionModel, ByVal index As Long, ByVal strain As Double, ByVal stress As Double)
    CrackAppendDiagnostic work, "bar=" & section.RebarID(index) & _
        "; strain=" & CrackFormatNumberInvariant(strain) & _
        "; stress=" & CrackFormatNumberInvariant(stress)
    If strain > work.MaxSteelStrain Then
        work.MaxSteelStrain = strain
        work.MaxSteelStress = stress
    End If
End Sub

' Возвращает phi3 для текущего LC. В Auto используется принятое правило СП:
' при растягивающей продольной силе принимается 1.2, иначе 1.0. В User
' программа не выводит характер нагрузки из НДС, а берет введенное значение.
Private Function ResolvePhi3(ByRef work As TCrackWidthWork, ByVal loadStateForClassification As CSectionLoadState) As Double
    If StrComp(work.Phi3Mode, "User", vbTextCompare) = 0 Then
        ResolvePhi3 = work.Phi3Input
    ElseIf IsTensionLoad(loadStateForClassification) Then
        ResolvePhi3 = CRACK_PHI3_TENSION
    Else
        ResolvePhi3 = CRACK_PHI3_BENDING
    End If
End Function

' Определяет знак продольной силы через общий объект нагрузки.
' Центральное растяжение и phi3 используют те же правила точки приложения
' и малых моментов, что прочность, несущая способность и устойчивость.
Private Function IsTensionLoad(ByVal loadStateForClassification As CSectionLoadState) As Boolean
    IsTensionLoad = loadStateForClassification.IsTension(CRACK_NEAR_ZERO_FORCE)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяет только завершение готовой формулы с другим пояснением события;
' НДС, коэффициенты, напряжения, расстояние ls и машинный признак не подменяются.
Friend Sub TestRefinalizeWithSpacingText(ByRef work As TCrackWidthWork, ByVal text As String)
    work.SpacingLimitWarning = text
    FinalizeWidth work, "Контроль представления готовой формулы."
End Sub
