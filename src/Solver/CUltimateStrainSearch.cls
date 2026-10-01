VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUltimateStrainSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Адаптер поиска UltimateStrain
' ==========================================================================
' Класс выполняет прямой поиск по предельной деформации и переиспользует общий
' Newton/line-search цикл для Capacity и CrackFormation. Если для силовой
' capacity-траектории прямой поиск не дал устойчивой точки, fallback на
' LoadMultiplier запускается через общий CLoadMultiplierSearch, без отдельной
' параллельной search-реализации.

' Выполняет ветку UltimateStrain для request-а несущей способности.
' Чисто осевая траектория переводится в LoadMultiplier тем же правилом, что и
' в основной capacity-логике.
Public Function ExecuteCapacity(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request Is Nothing Then Err.Raise vbObjectError + 4210, "CUltimateStrainSearch", "Не передан запрос limit-search."

    If request.PureAxialForMethod Then
        Set ExecuteCapacity = ExecuteCapacityLoadMultiplier(request)
        Exit Function
    End If

    Dim loadPath As CLoadPathVector
    Set loadPath = New CLoadPathVector
    loadPath.Initialize request.NOffset, request.NBase, request.MxOffset, request.MxBase, _
        request.MyOffset, request.MyBase
    RunCapacityUltimateNewton request, loadPath

    If request.ForceOnly And _
            StrComp(request.SolutionStrategy, "UltimateStrain", vbTextCompare) = 0 And _
            Not request.Problem.Converged And _
            Not request.Problem.StoppedTerminally Then
        Set ExecuteCapacity = ExecuteCapacityFallbackToLoadMultiplier(request, _
            "UltimateStrain не дал принятую предельную точку для силовой траектории", _
            "Переходим к LoadMultiplier по той же lambda-траектории.", _
            "UltimateStrain не сошелся, предельная точка найдена через резервный LoadMultiplier.")
        Exit Function
    End If

    Set ExecuteCapacity = request.Problem.BuildResult(request, request.Problem.Converged)
End Function

' Запускает ветку LoadMultiplier для capacity через общий search-класс.
' Используется для чисто осевых траекторий, где прямой UltimateStrain
' численно вырожден, и для fallback после неудачной force-only попытки.
Private Function ExecuteCapacityLoadMultiplier(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    Dim multiplier As CLoadMultiplierSearch
    Set multiplier = New CLoadMultiplierSearch
    Set ExecuteCapacityLoadMultiplier = multiplier.Execute(request)
End Function

' Повторяет ту же capacity-траекторию через LoadMultiplier и возвращает новый
' snapshot результата после объединения диагностики первой попытки и fallback.
Private Function ExecuteCapacityFallbackToLoadMultiplier(ByVal request As CLimitSearchRequest, _
        ByVal firstAttemptMessage As String, ByVal transitionMessage As String, _
        ByVal successMessage As String) As CLimitSearchResult
    Dim firstStopReason As String
    Dim firstLimitState As String
    Dim firstDiagnostics As String
    firstStopReason = request.Problem.StopReason
    firstLimitState = request.Problem.LimitState
    firstDiagnostics = request.Problem.DiagnosticLog

    Dim ignoredResult As CLimitSearchResult
    Set ignoredResult = ExecuteCapacityLoadMultiplier(request)
    request.Problem.MergeFallback firstLimitState, firstStopReason, _
        firstDiagnostics, firstAttemptMessage, transitionMessage, successMessage

    Set ExecuteCapacityFallbackToLoadMultiplier = request.Problem.BuildResult(request, request.Problem.Converged)
End Function

' Выполняет ветку UltimateStrain для поиска образования нормальной трещины.
' Низкоуровневый расчет остается в crack-domain контексте: там находятся
' критерий eps_bt,ult, сохранение PreCrackState/PostCrackState и диагностика.
Public Function ExecuteCrackFormation(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request Is Nothing Then Err.Raise vbObjectError + 4211, "CUltimateStrainSearch", "Не передан запрос limit-search."

    Dim searchSucceeded As Boolean
    searchSucceeded = request.Problem.RunUltimateAttempts(request)
    Set ExecuteCrackFormation = request.Problem.BuildResult(request, searchSucceeded)
End Function

' Выполняет общий Newton-поиск UltimateStrain для несущей способности.
' CCapacitySolver остается владельцем критерия, настроек и записи результата.
Public Function RunCapacityUltimateNewton(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector) As Boolean
    If Not request.Problem.PrepareUltimate(request, loadPath) Then Exit Function

    RunCapacityUltimateNewton = RunNewton(request.Problem, request.Section, _
        request.ConcreteMaterial, request.SteelMaterial, loadPath, Nothing, _
        "UltimateStrain: итерации Newton не сошлись.")
    request.Problem.CompleteUltimate loadPath
End Function

' Выполняет тот же общий Newton-поиск UltimateStrain для образования трещины.
' Crack-калькулятор передает выбранную lambda-траекторию и сохраняет результат.
Public Function RunCrackFormationNewton(ByVal problem As ILimitSearchProblem, _
        ByVal section As CSectionModel, ByVal concreteMaterial As CMaterialDiagram, _
        ByVal steelMaterial As CMaterialDiagram, ByVal loadPath As CLoadPathVector, _
        ByVal startSolver As CSectionSolver) As Boolean
    RunCrackFormationNewton = RunNewton(problem, section, concreteMaterial, _
        steelMaterial, loadPath, startSolver, _
        "NumericalFailure: прямой поиск lambda_crc по UltimateStrain не сошелся за заданное число итераций.")
End Function

' Общий цикл UltimateStrain: оценка невязок, численный Якобиан, Newton-шаг и
' line-search. Разные инженерные задачи отличаются только типизированными
' callback-ами; ни конкретный калькулятор, ни late-bound объект не требуются.
Public Function RunNewton(ByVal context As ILimitSearchProblem, ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByVal notConvergedReason As String) As Boolean
    If loadPath Is Nothing Then
        context.LimitSearchSetUltimateNumericalFailure "InputError: не задана lambda-траектория UltimateStrain."
        Exit Function
    End If

    Dim eps0 As Double
    Dim kx As Double
    Dim ky As Double
    context.LimitSearchInitializeUltimateGuess section, concreteMaterial, steelMaterial, _
        loadPath, startSolver, eps0, kx, ky

    Dim iter As Long
    Dim rPath1 As Double
    Dim rPath2 As Double
    Dim rLimit As Double
    Dim lambdaValue As Double
    Dim state As String
    Dim solver As CSectionSolver

    For iter = 1 To context.LimitSearchUltimateMaxIterations
        If Not context.LimitSearchEvaluateUltimateResidual(section, concreteMaterial, steelMaterial, _
                loadPath, eps0, kx, ky, solver, rPath1, rPath2, rLimit, lambdaValue, state) Then
            Exit Function
        End If

        If context.LimitSearchUltimateResidualReached(rPath1, rPath2, rLimit) Then
            If context.LimitSearchFinalizeUltimateResult(section, _
                    concreteMaterial, steelMaterial, loadPath, lambdaValue, solver, state, iter) Then
                RunNewton = True
                Exit Function
            End If
            If context.StoppedTerminally Then Exit Function
            ' Достижение внутренних residual-допусков не отменяет финальную
            ' проверку усилий: продолжаем уточнение той же физической точки.
        End If

        If Not ApplyUltimateNewtonStep(context, section, concreteMaterial, steelMaterial, _
                loadPath, eps0, kx, ky, rPath1, rPath2, rLimit) Then Exit Function
    Next iter

    context.LimitSearchSetUltimateNumericalFailure notConvergedReason
End Function

' Делает один общий Newton-шаг для UltimateStrain.
Private Function ApplyUltimateNewtonStep(ByVal context As ILimitSearchProblem, ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByRef eps0 As Double, _
        ByRef kx As Double, ByRef ky As Double, ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Boolean
    On Error GoTo Failed

    Dim hEps As Double
    Dim hKx As Double
    Dim hKy As Double
    hEps = MaxDouble(0.0000001, Abs(eps0) * 0.0001)
    hKx = MaxDouble(0.000000001, Abs(kx) * 0.0001)
    hKy = MaxDouble(0.000000001, Abs(ky) * 0.0001)
    ' У самого предела крупная разностная проба пересекает переломы диаграмм
    ' и дает производную другой активной ветви. Уменьшаем численный шаг
    ' по близости к критерию, сохраняя нижнюю границу от round-off.
    Dim differenceScale As Double
    differenceScale = MinDouble(1#, MaxDouble(0.001, Abs(rLimit)))
    hEps = hEps * differenceScale
    hKx = hKx * differenceScale
    hKy = hKy * differenceScale

    Dim j11 As Double, j21 As Double, j31 As Double
    Dim j12 As Double, j22 As Double, j32 As Double
    Dim j13 As Double, j23 As Double, j33 As Double
    BuildUltimateJacobianColumn context, section, concreteMaterial, steelMaterial, loadPath, _
        eps0 + hEps, kx, ky, hEps, rPath1, rPath2, rLimit, j11, j21, j31
    BuildUltimateJacobianColumn context, section, concreteMaterial, steelMaterial, loadPath, _
        eps0, kx + hKx, ky, hKx, rPath1, rPath2, rLimit, j12, j22, j32
    BuildUltimateJacobianColumn context, section, concreteMaterial, steelMaterial, loadPath, _
        eps0, kx, ky + hKy, hKy, rPath1, rPath2, rLimit, j13, j23, j33

    Dim system As CLinearSystem3x3
    Set system = New CLinearSystem3x3
    ' Поправки epsilon0 и кривизн имеют разные масштабы. Решаем ту же
    ' линейную систему в единицах шагов Якобиана и нормируем ее строки;
    ' допуски нелинейного равновесия и физического критерия не меняются.
    Dim rowScale1 As Double, rowScale2 As Double, rowScale3 As Double
    rowScale1 = MaxDouble(Abs(j11 * hEps), MaxDouble(Abs(j12 * hKx), Abs(j13 * hKy)))
    rowScale2 = MaxDouble(Abs(j21 * hEps), MaxDouble(Abs(j22 * hKx), Abs(j23 * hKy)))
    rowScale3 = MaxDouble(Abs(j31 * hEps), MaxDouble(Abs(j32 * hKx), Abs(j33 * hKy)))
    If rowScale1 <= 0# Then rowScale1 = 1#
    If rowScale2 <= 0# Then rowScale2 = 1#
    If rowScale3 <= 0# Then rowScale3 = 1#
    If Not system.Solve(j11 * hEps / rowScale1, j12 * hKx / rowScale1, j13 * hKy / rowScale1, _
            j21 * hEps / rowScale2, j22 * hKx / rowScale2, j23 * hKy / rowScale2, _
            j31 * hEps / rowScale3, j32 * hKx / rowScale3, j33 * hKy / rowScale3, _
            -rPath1 / rowScale1, -rPath2 / rowScale2, -rLimit / rowScale3) Then
        context.LimitSearchSetUltimateNumericalFailure _
            "NumericalFailure: матрица UltimateStrain вырождена: " & system.StopReason
        Exit Function
    End If

    Dim dEps As Double
    Dim dKx As Double
    Dim dKy As Double
    dEps = system.X1 * hEps
    dKx = system.X2 * hKx
    dKy = system.X3 * hKy
    LimitUltimateIncrement context, dEps, dKx, dKy

    Dim baseNorm As Double
    baseNorm = context.LimitSearchUltimateResidualNorm(rPath1, rPath2, rLimit)

    Dim alpha As Double
    alpha = context.LimitSearchUltimateDampingInitial
    If alpha <= 0# Or alpha > 1# Then alpha = 1#

    Do While alpha >= context.LimitSearchUltimateMinLineSearchAlpha
        Dim trialSolver As CSectionSolver
        Dim trialR1 As Double, trialR2 As Double, trialRLimit As Double
        Dim trialLambda As Double, trialState As String
        If context.LimitSearchEvaluateUltimateResidual(section, concreteMaterial, steelMaterial, _
                loadPath, eps0 + alpha * dEps, kx + alpha * dKx, ky + alpha * dKy, _
                trialSolver, trialR1, trialR2, trialRLimit, trialLambda, trialState) Then
            ' Уже достигнутые допуски важнее колебания нормы на уровне
            ' round-off. Итоговая приемка следующей итерации дополнительно
            ' подтверждает все усилия и физическую допустимость точки.
            If context.LimitSearchUltimateResidualReached(trialR1, trialR2, trialRLimit) Or _
                    (Not context.LimitSearchUltimateLineSearchEnabled) Or _
                    context.LimitSearchUltimateResidualNorm(trialR1, trialR2, trialRLimit) < baseNorm Then
                eps0 = eps0 + alpha * dEps
                kx = kx + alpha * dKx
                ky = ky + alpha * dKy
                ApplyUltimateNewtonStep = True
                Exit Function
            End If
        End If
        alpha = alpha / 2#
    Loop

    context.LimitSearchSetUltimateNumericalFailure _
        "NumericalFailure: line search UltimateStrain не улучшил невязку."
    Exit Function

Failed:
    context.LimitSearchSetUltimateNumericalFailure _
        "NumericalFailure: шаг Newton UltimateStrain не выполнен: " & Err.Description
End Function

' Строит один столбец численного Якобиана для общего UltimateStrain.
Private Sub BuildUltimateJacobianColumn(ByVal context As ILimitSearchProblem, ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal eps0 As Double, _
        ByVal kx As Double, ByVal ky As Double, ByVal stepValue As Double, _
        ByVal rPath1 As Double, ByVal rPath2 As Double, ByVal rLimit As Double, _
        ByRef j1 As Double, ByRef j2 As Double, ByRef j3 As Double)
    Dim solver As CSectionSolver
    Dim pR1 As Double, pR2 As Double, pLimit As Double
    Dim pLambda As Double, pState As String
    If Not context.LimitSearchEvaluateUltimateResidual(section, concreteMaterial, steelMaterial, _
            loadPath, eps0, kx, ky, solver, pR1, pR2, pLimit, pLambda, pState) Then Exit Sub
    j1 = (pR1 - rPath1) / stepValue
    j2 = (pR2 - rPath2) / stepValue
    j3 = (pLimit - rLimit) / stepValue
End Sub

' Ограничивает приращения плоскости деформаций по настройкам вызывающего расчета.
Private Sub LimitUltimateIncrement(ByVal context As ILimitSearchProblem, ByRef dEps As Double, _
        ByRef dKx As Double, ByRef dKy As Double)
    Dim scaleFactor As Double
    scaleFactor = 1#
    If context.LimitSearchUltimateMaxDeltaEpsilon0 > 0# And _
            Abs(dEps) > context.LimitSearchUltimateMaxDeltaEpsilon0 Then
        scaleFactor = MinDouble(scaleFactor, context.LimitSearchUltimateMaxDeltaEpsilon0 / Abs(dEps))
    End If
    If context.LimitSearchUltimateMaxDeltaKappa > 0# And _
            Abs(dKx) > context.LimitSearchUltimateMaxDeltaKappa Then
        scaleFactor = MinDouble(scaleFactor, context.LimitSearchUltimateMaxDeltaKappa / Abs(dKx))
    End If
    If context.LimitSearchUltimateMaxDeltaKappa > 0# And _
            Abs(dKy) > context.LimitSearchUltimateMaxDeltaKappa Then
        scaleFactor = MinDouble(scaleFactor, context.LimitSearchUltimateMaxDeltaKappa / Abs(dKy))
    End If
    dEps = dEps * scaleFactor
    dKx = dKx * scaleFactor
    dKy = dKy * scaleFactor
End Sub

' Утилитарный минимум без обращения к Excel.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

' Утилитарный максимум без обращения к Excel.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function
