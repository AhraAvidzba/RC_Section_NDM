VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCrackWidthSP35Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Раскрытие нормальных трещин по СП 35
' ==========================================================================
' По готовому НДС выполняет нормативную подготовку, формулы и выбор определяющего результата.
' Образование трещины и равновесие не пересчитывает; геометрия и диагностика общие.
' Состояние принадлежит вызывающему фасаду и не сохраняется в этом объекте.

' Готовит данные СП 35 по текущему НДС и выбирает наибольшую ширину
' среди пригодных областей взаимодействия. Образование трещины и НДС
' после неё не пересчитывает; коэффициент psi_s СП 63 здесь не применяется.
Friend Sub CalculateSP35(ByRef work As TCrackWidthWork, ByVal currentState As CSectionStateResult, ByVal section As CSectionModel, ByVal steelMaterial As Object)
    On Error GoTo Failed
    Dim gradient As Double, central As Boolean
    gradient = Sqr(currentState.KappaX * currentState.KappaX + currentState.KappaY * currentState.KappaY)
    central = work.CentralTensionBranch Or (gradient <= CRACK_NEAR_ZERO_STRAIN And CrackStrainAtConcreteCenter(currentState, section) > CRACK_NEAR_ZERO_STRAIN) ' Одинаковый с Prepare критерий по центру сечения сохраняет ветку при переносе WCS и при почти нулевой деформации.
    Set work.SP35Data = New CSP35CrackData
    work.SP35Data.Prepare section, currentState.Epsilon0, currentState.KappaX, currentState.KappaY, steelMaterial, _
        work.SP35GroupGapTolerance, work.SP35RowTolerance, work.SP35RadiusDiameterMode, work.SP35InteractionRadiusMode, central, _
        work.SP35NeighborRatioLimit, work.GeometryQuery
    work.CentralTensionBranch = central
    Dim i As Long, radius As Double, psiCm As Double, width As Double, region As CConcreteRegion
    For i = 1 To work.SP35Data.CandidateCount ' Геометрия и состав арматуры уже подготовлены; формулы получают площадь, реальные диаметры, напряжение и Es.
        If work.SP35Data.CandidateAvailable(i) Then
            Set region = work.SP35Data.CandidateRegion(i)
            radius = SP35ReinforcementRadiusFromData(region.Area, work.SP35Data.CandidateBetaDiameterSum(i))
            psiCm = SP35PsiFromData(radius, work.SP35RebarProfile)
            width = SP35CrackWidthFromData(work.SP35Data.CandidateSigmaS(i), work.SteelEs, psiCm)
            work.SP35Data.StoreFormulaResult i, radius, psiCm, width
            If work.SP35CriticalCandidate = 0 Or width > work.CrackWidth Then
                work.SP35CriticalCandidate = i: work.CrackWidth = width
                work.SP35Radius = radius: work.SP35PsiCm = psiCm
                Set work.InteractionRegion = region
            End If
            CrackAppendDiagnostic work, "СП 35: якорь G" & CStr(work.SP35Data.CandidateAnchorGroup(i)) & _
                "; Ar=" & CrackFormatNumberInvariant(region.Area) & "; Rr=" & CrackFormatNumberInvariant(radius) & _
                "; psi (см)=" & CrackFormatNumberInvariant(psiCm) & "; a_cr (мм)=" & CrackFormatNumberInvariant(width)
        End If
    Next i
    work.SP35Data.Freeze
    If work.SP35CriticalCandidate = 0 Then
        CrackFinishNoCrack work, "В текущем НДС нет пригодной растянутой арматуры и области взаимодействия для раскрытия по СП 35."
        Exit Sub
    End If

    work.Abt = work.InteractionRegion.Area: work.SigmaS = work.SP35Data.CandidateSigmaS(work.SP35CriticalCandidate) ' Сохраняем область с наибольшей шириной трещины и её фактическую арматуру.
    Dim rows As Variant, members As Variant, group As Long, j As Long, bar As Long
    rows = work.SP35Data.CandidateRows(work.SP35CriticalCandidate)
    For i = LBound(rows, 1) To UBound(rows, 1)
        group = CLng(rows(i, 1)): members = work.SP35Data.GroupMembers(group)
        For j = LBound(members) To UBound(members)
            bar = CLng(members(j))
            work.TensionRebarCount = work.TensionRebarCount + 1: work.RebarArea = work.RebarArea + work.SP35Data.BarArea(bar)
            If Len(work.TensionRebarIds) > 0 Then work.TensionRebarIds = work.TensionRebarIds & ", "
            work.TensionRebarIds = work.TensionRebarIds & work.SP35Data.BarID(bar)
        Next j
    Next i
    bar = work.SP35Data.CandidateStressBar(work.SP35CriticalCandidate)
    work.MaxSteelStrain = work.SP35Data.BarStrain(bar): work.MaxSteelStress = work.SigmaS
    work.FormulaCalculated = True: work.CrackFormed = True: work.Converged = True
    work.Utilization = CrackUtilizationFromData(work.CrackWidth, work.AllowableCrackWidth)
    Dim reason As String
    reason = "СП 35: раскрытие проверено по текущему НДС; критическая группа G" & _
        CStr(work.SP35Data.CandidateAnchorGroup(work.SP35CriticalCandidate)) & ", напряжение стержня " & work.SP35Data.BarID(bar) & "."
    If work.CrackWidth > work.AllowableCrackWidth Then
        CrackSetWidthResult work, rsCheckFailed, rcCrackWidthExceeded, _
            "Раскрытие нормальной трещины " & CrackFormatNumberInvariant(work.CrackWidth) & " мм превышает допускаемое " & _
            CrackFormatNumberInvariant(work.AllowableCrackWidth) & " мм. " & reason
    Else
        CrackSetWidthResult work, rsSuccess, rcCheckPassed, reason
    End If
    Exit Sub
Failed:
    Dim failureNumber As Long, failureReason As String
    failureNumber = Err.Number: failureReason = Err.Description
    work.Converged = False: work.FormulaCalculated = False: work.CrackFormed = False
    Set work.InteractionRegion = Nothing: Set work.SP35Data = Nothing
    work.SP35CriticalCandidate = 0: work.CrackWidth = 0#: work.Utilization = 0#
    Select Case failureNumber
        Case vbObjectError + 5510, vbObjectError + 5401, vbObjectError + 3780 To vbObjectError + 3785
            CrackSetWidthResult work, rsInvalidInput, rcInvalidInput, "Для раскрытия по СП 35: " & failureReason
        Case Else
            CrackSetWidthResult work, rsInternalError, rcInternalError, "Не удалось подготовить раскрытие по СП 35: " & failureReason
    End Select
End Sub

' ------------------------------ ЧИСТЫЕ ФОРМУЛЫ СП 35 ------------------------------
' [SP35] П. 7.110, пособие (4.310): вычисляет радиус армирования, мм.
' Область взаимодействия и состав групп подготовлены заранее;
' формула получает числа и не выполняет геометрический поиск или решение НДС.
Friend Function SP35ReinforcementRadiusFromData(ByVal interactionArea As Double, ByVal betaDiameterSum As Double) As Double
    If interactionArea < 0# Or betaDiameterSum <= 0# Then _
        Err.Raise vbObjectError + 3780, "CCrackWidthCalculator", _
            "Для радиуса армирования площадь должна быть неотрицательной, а сумма диаметров с коэффициентами сцепления - положительной."
    SP35ReinforcementRadiusFromData = interactionArea / betaDiameterSum
End Function

' [SP35] П. 7.109, пособие (4.308)-(4.309): psi имеет размерность см.
' Нормативные коэффициенты используют Rr в см; входной радиус хранится в мм.
' Строка профиля уже проверена слоем настроек, неизвестный вариант не заменяется.
Friend Function SP35PsiFromData(ByVal reinforcementRadiusMm As Double, ByVal rebarProfile As String) As Double
    If reinforcementRadiusMm < 0# Then Err.Raise vbObjectError + 3781, "CCrackWidthCalculator", "Радиус армирования не может быть отрицательным."
    Dim radiusCm As Double
    radiusCm = reinforcementRadiusMm / 10# ' Коэффициенты формулы СП 35 используют сантиметры; основная геометрия хранится в миллиметрах.
    Select Case LCase$(Trim$(rebarProfile))
        Case "periodic": SP35PsiFromData = 1.5 * Sqr(radiusCm)
        Case "smooth": SP35PsiFromData = 0.35 * radiusCm
        Case Else
            Err.Raise vbObjectError + 3782, "CCrackWidthCalculator", "Для формулы СП 35 профиль арматуры должен быть Periodic или Smooth."
    End Select
End Function

' [SP35] П. 7.105, пособие (4.306): a_cr=(sigma_s/Es)*psi.
' Напряжение и модуль задаются в МПа, psi — в см; результат возвращается в мм.
' До вызова выбираются наибольшее напряжение проверяемой группы и соответствующий модуль.
Friend Function SP35CrackWidthFromData(ByVal sigmaS As Double, ByVal steelEs As Double, ByVal psiCm As Double) As Double
    If sigmaS < 0# Or steelEs <= 0# Or psiCm < 0# Then _
        Err.Raise vbObjectError + 3783, "CCrackWidthCalculator", _
            "Для формулы СП 35 напряжение и коэффициент раскрытия должны быть неотрицательными, а модуль арматуры - положительным."
    SP35CrackWidthFromData = (sigmaS / steelEs) * psiCm * 10#
End Function
