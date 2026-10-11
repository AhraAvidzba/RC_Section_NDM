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
' Общий поиск предельной плоскости UltimateStrain
' ==========================================================================
' Класс выполняет прямой поиск по предельной деформации и переиспользует общий
' Newton/line-search цикл для Capacity и CrackFormation. Если для силовой
' capacity-траектории прямой поиск не дал устойчивой точки, fallback на
' LoadMultiplier запускается через общий CLoadMultiplierSearch, без отдельной
' параллельной search-реализации.

Private mFrame As CSectionNumericalFrame ' Тот же центр и масштаб, что у прямого НДС.
Private mMeritWeight1 As Double, mMeritWeight2 As Double, mMeritWeight3 As Double
Private mMeritScalingReady As Boolean ' Веса текущего шага по реальным строкам Якобиана; не допуски приема точки.

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
    Dim firstSearchExecuted As Boolean
    firstSearchExecuted = request.Problem.BuildResult(request, False).SearchExecuted

    Dim ignoredResult As CLimitSearchResult
    Set ignoredResult = ExecuteCapacityLoadMultiplier(request)
    request.Problem.MergeFallback firstLimitState, firstStopReason, _
        firstDiagnostics, firstAttemptMessage, transitionMessage, successMessage, firstSearchExecuted

    Set ExecuteCapacityFallbackToLoadMultiplier = request.Problem.BuildResult(request, request.Problem.Converged)
End Function

' Выполняет ветку UltimateStrain для поиска образования нормальной трещины.
' Инженерный критерий eps_bt,ult, сохранение PreCrackState/PostCrackState и
' диагностика принадлежат Formation; численный Newton-цикл общий с Capacity.
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

    RunCapacityUltimateNewton = request.Problem.RunUltimateAttempts(request)
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
        "Поиск точки образования трещины по предельной деформации не сошелся за заданное число итераций.")
End Function

' Общий цикл UltimateStrain: оценка невязок, численный Якобиан, Newton-шаг и
' line-search. Разные инженерные задачи отличаются только типизированными
' callback-ами; ни конкретный калькулятор, ни late-bound объект не требуются.
Public Function RunNewton(ByVal context As ILimitSearchProblem, ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByVal notConvergedReason As String) As Boolean
    If context Is Nothing Then Err.Raise vbObjectError + 4212, "CUltimateStrainSearch", _
        "Не передана задача общего поиска UltimateStrain."
    On Error GoTo Failed
    If loadPath Is Nothing Then
        context.SetFailure sfcInternalError, "Общий поиск UltimateStrain не получил траекторию нагрузки."
        Exit Function
    End If
    If Not ValidateNewtonSettings(context) Then Exit Function
    mMeritScalingReady = False

    Set mFrame = New CSectionNumericalFrame
    If Not section Is Nothing Then
        mFrame.Initialize section
    End If

    Dim eps0 As Double
    Dim kx As Double
    Dim ky As Double
    context.LimitSearchInitializeUltimateGuess section, concreteMaterial, steelMaterial, _
        loadPath, startSolver, eps0, kx, ky
    ' Newton меняет деформацию в центре, а не в удаленном начале чертежа.
    ' В доменный callback передается та же физическая плоскость в исходных осях.
    eps0 = mFrame.CenterStrain(eps0, kx, ky)
    If Not IsFiniteNumber(eps0) Or Not IsFiniteNumber(kx) Or Not IsFiniteNumber(ky) Then
        context.SetFailure sfcNumericalFailure, "Стартовая плоскость UltimateStrain содержит нечисловое или бесконечное значение."
        Exit Function
    End If

    Dim iter As Long
    Dim rPath1 As Double
    Dim rPath2 As Double
    Dim rLimit As Double
    Dim lambdaValue As Double
    Dim state As String
    Dim solver As CSectionSolver
    Dim acceptedEvaluation As Boolean, acceptedRevision As Long, immutableMaterials As Boolean
    Dim acceptedPath(1 To 6) As Double ' Точный Offset/Base; без строковых или округленных ключей.
    If Not section Is Nothing And Not concreteMaterial Is Nothing And Not steelMaterial Is Nothing Then
        If TypeOf concreteMaterial Is CMaterialDiagram And TypeOf steelMaterial Is CMaterialDiagram Then
            Dim concreteDiagram As CMaterialDiagram, steelDiagram As CMaterialDiagram
            Set concreteDiagram = concreteMaterial: Set steelDiagram = steelMaterial
            immutableMaterials = concreteDiagram.IsFrozen And steelDiagram.IsFrozen
        End If
    End If

    For iter = 1 To context.LimitSearchUltimateMaxIterations
        If acceptedEvaluation Then
            acceptedEvaluation = section.Revision = acceptedRevision And _
                loadPath.NOffset = acceptedPath(1) And loadPath.NBase = acceptedPath(2) And _
                loadPath.MxOffset = acceptedPath(3) And loadPath.MxBase = acceptedPath(4) And _
                loadPath.MyOffset = acceptedPath(5) And loadPath.MyBase = acceptedPath(6)
        End If
        If Not acceptedEvaluation Then
            If Not EvaluateResidual(context, section, concreteMaterial, steelMaterial, _
                    loadPath, eps0, kx, ky, solver, rPath1, rPath2, rLimit, lambdaValue, state) Then
                Exit Function
            End If
        End If

        If context.LimitSearchUltimateResidualReached(rPath1, rPath2, rLimit) Then
            Dim finalization As EUltimateFinalization
            finalization = context.LimitSearchFinalizeUltimateResult(section, _
                concreteMaterial, steelMaterial, loadPath, lambdaValue, solver, state, iter)
            Select Case finalization
                Case ufAccepted
                    RunNewton = True
                    Exit Function
                Case ufRejectedPoint
                    Exit Function
                Case ufNeedsRefinement
                Case Else
                    context.SetFailure sfcInternalError, "Неизвестный исход приемки предельной точки UltimateStrain."
                    Exit Function
            End Select
            If context.StoppedTerminally Then Exit Function
            ' Достижение внутренних residual-допусков не отменяет финальную
            ' проверку усилий: продолжаем уточнение той же физической точки.
        End If

        context.AppendDiagnostic "ultimateNewton step=" & CStr(iter)
        If Not ApplyUltimateNewtonStep(context, section, concreteMaterial, steelMaterial, _
                loadPath, eps0, kx, ky, rPath1, rPath2, rLimit, solver, lambdaValue, state) Then Exit Function
        ' Принятая проба уже имеет свои residual/solver/критерий. Пробы Якобиана
        ' ее не подменяют; финальное подтверждение усилий и физики остается выше.
        acceptedEvaluation = immutableMaterials
        If acceptedEvaluation Then
            acceptedRevision = section.Revision
            acceptedPath(1) = loadPath.NOffset: acceptedPath(2) = loadPath.NBase
            acceptedPath(3) = loadPath.MxOffset: acceptedPath(4) = loadPath.MxBase
            acceptedPath(5) = loadPath.MyOffset: acceptedPath(6) = loadPath.MyBase
        End If
    Next iter

    ' Последняя разрешенная поправка тоже проходит обычную финальную приемку.
    ' Новая итерация не выполняется, бюджет поиска не увеличивается.
    If Not EvaluateResidual(context, section, concreteMaterial, steelMaterial, _
            loadPath, eps0, kx, ky, solver, rPath1, rPath2, rLimit, lambdaValue, state) Then Exit Function
    If context.LimitSearchUltimateResidualReached(rPath1, rPath2, rLimit) Then
        finalization = context.LimitSearchFinalizeUltimateResult(section, concreteMaterial, steelMaterial, _
            loadPath, lambdaValue, solver, state, context.LimitSearchUltimateMaxIterations)
        Select Case finalization
            Case ufAccepted
                RunNewton = True
                Exit Function
            Case ufRejectedPoint
                Exit Function
            Case ufNeedsRefinement
            Case Else
                context.SetFailure sfcInternalError, "Неизвестный исход приемки предельной точки UltimateStrain."
                Exit Function
        End Select
        If context.StoppedTerminally Then Exit Function
    End If
    context.LimitSearchSetUltimateNumericalFailure notConvergedReason
    Exit Function
Failed:
    ReportArithmeticOrContractFailure context, "Общий поиск UltimateStrain", Err.Number, Err.Description
End Function

' Пользователь задает бюджет итераций; поиск длины шага автоматический.
Private Function ValidateNewtonSettings(ByVal context As ILimitSearchProblem) As Boolean
    If context.LimitSearchUltimateMaxIterations <= 0 Then
        context.SetFailure sfcInvalidConfiguration, "Число итераций UltimateStrain должно быть положительным."
    Else
        ValidateNewtonSettings = True
    End If
End Function

' Делает один общий Newton-шаг для UltimateStrain.
Private Function ApplyUltimateNewtonStep(ByVal context As ILimitSearchProblem, ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByRef eps0 As Double, _
        ByRef kx As Double, ByRef ky As Double, ByRef rPath1 As Double, _
        ByRef rPath2 As Double, ByRef rLimit As Double, ByRef solver As CSectionSolver, _
        ByRef lambdaValue As Double, ByRef state As String) As Boolean
    On Error GoTo Failed

    Dim hEps As Double
    Dim hKx As Double
    Dim hKy As Double
    hEps = MaxDouble(0.0000001, Abs(eps0) * 0.0001)
    hKx = MaxDouble(0.0000001 / mFrame.LengthScale, Abs(kx) * 0.0001)
    hKy = MaxDouble(0.0000001 / mFrame.LengthScale, Abs(ky) * 0.0001)
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
    If eps0 + hEps = eps0 Or kx + hKx = kx Or ky + hKy = ky Then
        context.SetFailure sfcNumericalFailure, "Шаг численного Якобиана UltimateStrain не представим при текущем масштабе плоскости деформаций."
        Exit Function
    End If
    If Not BuildUltimateJacobianColumn(context, section, concreteMaterial, steelMaterial, loadPath, _
        eps0 + hEps, kx, ky, hEps, rPath1, rPath2, rLimit, j11, j21, j31) Then Exit Function
    If Not BuildUltimateJacobianColumn(context, section, concreteMaterial, steelMaterial, loadPath, _
        eps0, kx + hKx, ky, hKx, rPath1, rPath2, rLimit, j12, j22, j32) Then Exit Function
    If Not BuildUltimateJacobianColumn(context, section, concreteMaterial, steelMaterial, loadPath, _
        eps0, kx, ky + hKy, hKy, rPath1, rPath2, rLimit, j13, j23, j33) Then Exit Function

    Dim system As CLinearSystem3x3
    Set system = New CLinearSystem3x3
    ' Поправки epsilon0 и кривизн имеют разные масштабы. Решаем ту же
    ' линейную систему в единицах шагов Якобиана и нормируем ее строки;
    ' допуски нелинейного равновесия и физического критерия не меняются.
    Dim rowScale1 As Double, rowScale2 As Double, rowScale3 As Double
    rowScale1 = MaxDouble(Abs(j11 * hEps), MaxDouble(Abs(j12 * hKx), Abs(j13 * hKy)))
    rowScale2 = MaxDouble(Abs(j21 * hEps), MaxDouble(Abs(j22 * hKx), Abs(j23 * hKy)))
    rowScale3 = MaxDouble(Abs(j31 * hEps), MaxDouble(Abs(j32 * hKx), Abs(j33 * hKy)))
    If context.LimitSearchUseJacobianMeritScaling Then
        ConfigureStepResidualScales rowScale1, rowScale2, rowScale3
    Else
        mMeritScalingReady = False
    End If
    If rowScale1 <= 0# Then rowScale1 = 1#
    If rowScale2 <= 0# Then rowScale2 = 1#
    If rowScale3 <= 0# Then rowScale3 = 1#
    Dim a(1 To 3, 1 To 3) As Double, b(1 To 3) As Double
    a(1, 1) = j11 * hEps / rowScale1: a(1, 2) = j12 * hKx / rowScale1: a(1, 3) = j13 * hKy / rowScale1
    a(2, 1) = j21 * hEps / rowScale2: a(2, 2) = j22 * hKx / rowScale2: a(2, 3) = j23 * hKy / rowScale2
    a(3, 1) = j31 * hEps / rowScale3: a(3, 2) = j32 * hKx / rowScale3: a(3, 3) = j33 * hKy / rowScale3
    b(1) = -rPath1 / rowScale1: b(2) = -rPath2 / rowScale2: b(3) = -rLimit / rowScale3
    Dim ordinarySolved As Boolean, linearFailure As ESolverFailureCode, linearReason As String
    ordinarySolved = system.Solve(a(1, 1), a(1, 2), a(1, 3), a(2, 1), a(2, 2), a(2, 3), _
        a(3, 1), a(3, 2), a(3, 3), b(1), b(2), b(3))
    linearFailure = system.FailureCode: linearReason = system.StopReason
    If ordinarySolved Then
        If TryUltimateDirection(context, section, concreteMaterial, steelMaterial, loadPath, _
                eps0, kx, ky, rPath1, rPath2, rLimit, solver, lambdaValue, state, _
                system.X1 * hEps, system.X2 * hKx, system.X3 * hKy) Then
            ApplyUltimateNewtonStep = True
            Exit Function
        End If
        If context.StoppedTerminally Then Exit Function
    End If

    ' Локальная производная на плато не видит возврата на активную ветвь.
    ' Двусторонние секущие последовательно расширяют пробу до масштаба
    ' текущей плоскости; результат по-прежнему принимает только line-search.
    If mFrame.HasFullPlaneSupport Then
        Dim horizon As Double, sample As Long, sampleStep As Double
        horizon = 2# * MaxDouble(0.0000001, MaxDouble(Abs(eps0), _
            mFrame.LengthScale * MaxDouble(Abs(kx), Abs(ky))))
        For sample = 0 To 6
            sampleStep = horizon * 10# ^ (sample - 6)
            If TryUltimateProjectedStep(context, section, concreteMaterial, steelMaterial, loadPath, _
                    eps0, kx, ky, rPath1, rPath2, rLimit, solver, lambdaValue, state, _
                    MaxDouble(hEps, sampleStep), j31, j32, j33, horizon) Then
                ApplyUltimateNewtonStep = True
                Exit Function
            End If
            If context.StoppedTerminally Then Exit Function
            If sample = 0 Then
                If TryUltimateMaterialStep(context, section, concreteMaterial, steelMaterial, loadPath, _
                        eps0, kx, ky, rPath1, rPath2, rLimit, solver, lambdaValue, state, _
                        hEps, j31, j32, j33) Then
                    ApplyUltimateNewtonStep = True
                    Exit Function
                End If
                If context.StoppedTerminally Then Exit Function
            End If
        Next sample
    End If
    If Not ordinarySolved Then
        context.SetFailure linearFailure, "Не удалось решить линейную систему UltimateStrain: " & linearReason
    Else
        context.LimitSearchSetUltimateNumericalFailure _
            "Поиск по предельной деформации не сошелся: подбор шага не уменьшил невязку."
    End If
    Exit Function
Failed:
    ReportArithmeticOrContractFailure context, "Шаг Newton UltimateStrain", Err.Number, Err.Description
End Function

' Две секущие строятся в плоскости активного деформационного ограничения.
' Поэтому проба усилий не смешивается с независимым уходом от этого предела.
Private Function TryUltimateProjectedStep(ByVal context As ILimitSearchProblem, ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, ByVal loadPath As CLoadPathVector, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double, ByRef rPath1 As Double, _
        ByRef rPath2 As Double, ByRef rLimit As Double, ByRef solver As CSectionSolver, _
        ByRef lambdaValue As Double, ByRef state As String, ByVal h As Double, _
        ByVal limitEps As Double, ByVal limitKx As Double, ByVal limitKy As Double, _
        ByVal horizon As Double) As Boolean
    Dim q(1 To 3) As Double, correction(1 To 3) As Double
    Dim v(1 To 3, 1 To 2) As Double, i As Long, j As Long
    If Not PrepareUltimateFace(eps0, kx, ky, rLimit, limitEps, limitKx, limitKy, q, correction, v) Then Exit Function
    Dim trial As CSectionSolver, p1 As Double, p2 As Double, p3 As Double
    Dim trialLambda As Double, trialState As String
    If Not EvaluateResidual(context, section, concreteMaterial, steelMaterial, loadPath, _
            q(1), q(2) / mFrame.LengthScale, q(3) / mFrame.LengthScale, _
            trial, p1, p2, p3, trialLambda, trialState) Then Exit Function
    ' Нормальная поправка сама может достигнуть предела на плато.
    ' Ее принимают только реальные невязки задачи, не локальная модель шага.
    If context.LimitSearchUltimateResidualReached(p1, p2, p3) Then
        eps0 = q(1): kx = q(2) / mFrame.LengthScale: ky = q(3) / mFrame.LengthScale
        rPath1 = p1: rPath2 = p2: rLimit = p3
        Set solver = trial: lambdaValue = trialLambda: state = trialState
        TryUltimateProjectedStep = True
        Exit Function
    End If
    Dim a(1 To 2, 1 To 2) As Double, b(1 To 2) As Double, matrixMagnitude As Double
    For j = 1 To 2
        If Not BuildUltimateProjectedColumn(context, section, concreteMaterial, steelMaterial, loadPath, _
                q, v, j, h, a(1, j), a(2, j)) Then Exit Function
    Next j
    b(1) = -p1: b(2) = -p2
    For i = 1 To 2
        matrixMagnitude = MaxDouble(matrixMagnitude, MaxDouble(Abs(a(i, 1)), Abs(a(i, 2))))
    Next i
    If matrixMagnitude = 0# Then Exit Function
    ' Общий масштаб сохраняет веса невязок. Независимое нормирование строк
    ' изменило бы минимум резервного least-squares шага на вырожденной матрице.
    For i = 1 To 2
        a(i, 1) = a(i, 1) / matrixMagnitude: a(i, 2) = a(i, 2) / matrixMagnitude
        b(i) = b(i) / matrixMagnitude
    Next i
    ' Dogleg ограничивает шаг по качеству локальной модели. При отказе
    ' сохраняются прежние Newton/least-squares направления и их line-search.
    If TryProjectedTrustRegion(context, section, concreteMaterial, steelMaterial, loadPath, _
            eps0, kx, ky, rPath1, rPath2, rLimit, solver, lambdaValue, state, _
            q, v, a, b, matrixMagnitude, h, horizon) Then
        TryUltimateProjectedStep = True
        Exit Function
    End If
    If context.StoppedTerminally Then Exit Function
    Dim system As CLinearSystem3x3
    Set system = New CLinearSystem3x3
    If system.Solve(a(1, 1), a(1, 2), 0#, a(2, 1), a(2, 2), 0#, 0#, 0#, 1#, b(1), b(2), 0#) Then
        If TryUltimateProjectedDirection(context, section, concreteMaterial, steelMaterial, loadPath, _
                eps0, kx, ky, rPath1, rPath2, rLimit, solver, lambdaValue, state, _
                correction, v, h * system.X1, h * system.X2, horizon) Then
            TryUltimateProjectedStep = True
            Exit Function
        End If
        If context.StoppedTerminally Then Exit Function
    End If
    Dim normal(1 To 2, 1 To 2) As Double, gradient(1 To 2) As Double, k As Long
    For i = 1 To 2
        For k = 1 To 2
            gradient(i) = gradient(i) + a(k, i) * b(k)
            For j = 1 To 2
                normal(i, j) = normal(i, j) + a(k, i) * a(k, j)
            Next j
        Next k
    Next i
    If gradient(1) = 0# And gradient(2) = 0# Then Exit Function
    Dim mu As Double, shift As Long
    mu = 0.00000001
    For shift = 0 To 10
        If system.Solve(normal(1, 1) + mu, normal(1, 2), 0#, normal(2, 1), normal(2, 2) + mu, _
                0#, 0#, 0#, 1#, gradient(1), gradient(2), 0#) Then
            If TryUltimateProjectedDirection(context, section, concreteMaterial, steelMaterial, loadPath, _
                    eps0, kx, ky, rPath1, rPath2, rLimit, solver, lambdaValue, state, _
                    correction, v, h * system.X1, h * system.X2, horizon) Then
                TryUltimateProjectedStep = True
                Exit Function
            End If
            If context.StoppedTerminally Then Exit Function
        End If
        mu = mu * 10#
    Next shift
End Function

' Два касательных направления активного предела: Newton внутри радиуса,
' Cauchy или соединяющий их dogleg на границе. Радиус уменьшается при
' плохом прогнозе; допуски и окончательная приемка принадлежат задаче.
Private Function TryProjectedTrustRegion(ByVal context As ILimitSearchProblem, ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, ByVal loadPath As CLoadPathVector, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double, ByRef rPath1 As Double, _
        ByRef rPath2 As Double, ByRef rLimit As Double, ByRef solver As CSectionSolver, _
        ByRef lambdaValue As Double, ByRef state As String, ByRef q() As Double, _
        ByRef v() As Double, ByRef a() As Double, ByRef b() As Double, ByVal matrixMagnitude As Double, _
        ByVal h As Double, ByVal horizon As Double) As Boolean
    Dim system As CLinearSystem3x3, hasNewton As Boolean, newton1 As Double, newton2 As Double
    Set system = New CLinearSystem3x3
    hasNewton = system.Solve(a(1, 1), a(1, 2), 0#, a(2, 1), a(2, 2), 0#, 0#, 0#, 1#, b(1), b(2), 0#)
    If hasNewton Then newton1 = system.X1: newton2 = system.X2
    Dim g1 As Double, g2 As Double, ag1 As Double, ag2 As Double, gNorm As Double, agNorm As Double
    g1 = a(1, 1) * b(1) + a(2, 1) * b(2)
    g2 = a(1, 2) * b(1) + a(2, 2) * b(2)
    gNorm = ProjectedVectorNorm(g1, g2)
    If gNorm = 0# Then Exit Function
    ag1 = a(1, 1) * (g1 / gNorm) + a(1, 2) * (g2 / gNorm)
    ag2 = a(2, 1) * (g1 / gNorm) + a(2, 2) * (g2 / gNorm)
    agNorm = ProjectedVectorNorm(ag1, ag2)
    If agNorm = 0# Then Exit Function
    Dim cauchyLength As Double, cauchy1 As Double, cauchy2 As Double
    cauchyLength = (gNorm / agNorm) / agNorm
    cauchy1 = (g1 / gNorm) * cauchyLength: cauchy2 = (g2 / gNorm) * cauchyLength
    Dim fullLength As Double, radius As Double, baseNorm As Double, attempt As Long
    If hasNewton Then fullLength = ProjectedVectorNorm(newton1, newton2) Else fullLength = cauchyLength
    radius = horizon / h
    If radius <= 0# Then Exit Function
    baseNorm = StepResidualNorm(context, rPath1, rPath2, rLimit)
    For attempt = 1 To NONLINEAR_MAX_BACKTRACKS
        Dim s1 As Double, s2 As Double, stepLength As Double, alpha As Double
        ProjectedDogleg g1 / gNorm, g2 / gNorm, cauchy1, cauchy2, cauchyLength, _
            hasNewton, newton1, newton2, fullLength, radius, s1, s2
        stepLength = ProjectedVectorNorm(s1, s2)
        If stepLength = 0# Then Exit Function
        Dim ap1 As Double, ap2 As Double, predicted As Double
        ap1 = a(1, 1) * s1 + a(1, 2) * s2
        ap2 = a(2, 1) * s1 + a(2, 2) * s2
        predicted = 2# * (g1 * s1 + g2 * s2) - ap1 * ap1 - ap2 * ap2
        If predicted <= 0# Then Exit Function
        Dim trialE As Double, trialX As Double, trialY As Double
        trialE = q(1) + h * (v(1, 1) * s1 + v(1, 2) * s2)
        trialX = (q(2) + h * (v(2, 1) * s1 + v(2, 2) * s2)) / mFrame.LengthScale
        trialY = (q(3) + h * (v(3, 1) * s1 + v(3, 2) * s2)) / mFrame.LengthScale
        If trialE = eps0 And trialX = kx And trialY = ky Then Exit Function
        Dim trial As CSectionSolver, t1 As Double, t2 As Double, t3 As Double
        Dim trialLambda As Double, trialState As String, rho As Double, actual As Double
        If Not EvaluateResidual(context, section, concreteMaterial, steelMaterial, loadPath, _
                trialE, trialX, trialY, trial, t1, t2, t3, trialLambda, trialState) Then Exit Function
        ' Разность квадратов без вычитания двух огромных близких норм.
        actual = (b(1) - t1 / matrixMagnitude) * (b(1) + t1 / matrixMagnitude) + _
            (b(2) - t2 / matrixMagnitude) * (b(2) + t2 / matrixMagnitude)
        rho = actual / predicted
        alpha = 1#
        If fullLength > stepLength Then alpha = stepLength / fullLength
        If context.LimitSearchUltimateResidualReached(t1, t2, t3) Or _
                (rho >= 0.001 And NonlinearStepImproves(baseNorm, StepResidualNorm(context, t1, t2, t3), alpha)) Then
            If Not context.LimitSearchUltimateResidualReached(t1, t2, t3) Then
                If ProjectedMaterialsAreFlat(section, concreteMaterial, steelMaterial, trialE, trialX, trialY) Then
                    context.AppendDiagnostic "ultimateTrustRegion flat trial rejected; original recovery remains available"
                    Exit Function
                End If
            End If
            eps0 = trialE: kx = trialX: ky = trialY
            rPath1 = t1: rPath2 = t2: rLimit = t3
            Set solver = trial: lambdaValue = trialLambda: state = trialState
            context.AppendDiagnostic "ultimateTrustRegion attempt=" & CStr(attempt) & "; rho=" & CStr(rho)
            TryProjectedTrustRegion = True
            Exit Function
        End If
        If context.StoppedTerminally Then Exit Function
        radius = 0.25 * MinDouble(radius, stepLength)
    Next attempt
End Function

' Улучшившая норму, но полностью горизонтальная проба может запереть
' следующий шаг на плато. Тогда оставляем прежнее восстановление; достигнутое
' равновесие выше принимается независимо от касательной материалов.
Private Function ProjectedMaterialsAreFlat(ByVal section As CSectionModel, ByVal concreteMaterial As Object, _
        ByVal steelMaterial As Object, ByVal epsCenter As Double, ByVal kx As Double, ByVal ky As Double) As Boolean
    If Not TypeOf concreteMaterial Is CMaterialDiagram Then Exit Function
    If Not TypeOf steelMaterial Is CMaterialDiagram Then Exit Function
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, strain As Double, i As Long
    Set concrete = concreteMaterial: Set steel = steelMaterial
    For i = 1 To section.ConcreteCount
        strain = epsCenter + kx * (section.ConcreteY(i) - mFrame.OriginY) + ky * (section.ConcreteX(i) - mFrame.OriginX)
        If concrete.GetTangentModulus(strain) <> 0# Then Exit Function
    Next i
    For i = 1 To section.RebarCount
        strain = epsCenter + kx * (section.RebarY(i) - mFrame.OriginY) + ky * (section.RebarX(i) - mFrame.OriginX)
        If steel.GetTangentModulus(strain) <> 0# Then Exit Function
        If concrete.GetTangentModulus(strain) <> 0# Then Exit Function
    Next i
    ProjectedMaterialsAreFlat = True
End Function

' Устойчивая длина в масштабированных координатах без квадратов больших чисел.
Private Function ProjectedVectorNorm(ByVal x As Double, ByVal y As Double) As Double
    Dim magnitude As Double
    magnitude = MaxDouble(Abs(x), Abs(y))
    If magnitude > 0# Then ProjectedVectorNorm = magnitude * Sqr((x / magnitude) ^ 2 + (y / magnitude) ^ 2)
End Function

Private Sub ProjectedDogleg(ByVal gradient1 As Double, ByVal gradient2 As Double, _
        ByVal c1 As Double, ByVal c2 As Double, ByVal cLength As Double, _
        ByVal hasNewton As Boolean, ByVal n1 As Double, ByVal n2 As Double, ByVal nLength As Double, _
        ByVal radius As Double, ByRef s1 As Double, ByRef s2 As Double)
    If hasNewton And nLength <= radius Then
        s1 = n1: s2 = n2
    ElseIf cLength >= radius Then
        s1 = gradient1 * radius: s2 = gradient2 * radius
    ElseIf hasNewton Then
        Dim d1 As Double, d2 As Double, size As Double, aa As Double, dot As Double, gap As Double, tau As Double
        size = MaxDouble(radius, nLength)
        d1 = (n1 - c1) / size: d2 = (n2 - c2) / size
        aa = d1 * d1 + d2 * d2
        If aa = 0# Then
            s1 = c1: s2 = c2
            Exit Sub
        End If
        dot = (c1 / size) * d1 + (c2 / size) * d2
        gap = (radius / size) ^ 2 - (cLength / size) ^ 2
        tau = (-dot + Sqr(dot * dot + aa * gap)) / aa
        s1 = c1 + tau * (n1 - c1): s2 = c2 + tau * (n2 - c2)
    Else
        s1 = c1: s2 = c2
    End If
End Sub

' Расширяется только касательная часть шага. Нормальная поправка к пределу
' не умножается вместе с ней; финальная проба проверяет все реальные невязки.
Private Function TryUltimateProjectedDirection(ByVal context As ILimitSearchProblem, ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, ByVal loadPath As CLoadPathVector, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double, ByRef rPath1 As Double, _
        ByRef rPath2 As Double, ByRef rLimit As Double, ByRef solver As CSectionSolver, _
        ByRef lambdaValue As Double, ByRef state As String, ByRef correction() As Double, _
        ByRef v() As Double, ByVal t1 As Double, ByVal t2 As Double, ByVal horizon As Double) As Boolean
    Dim tangent(1 To 3) As Double, i As Long, size As Double, factor As Double, attempt As Long
    For i = 1 To 3
        tangent(i) = v(i, 1) * t1 + v(i, 2) * t2
        size = MaxDouble(size, Abs(tangent(i)))
    Next i
    factor = 1#
    If size > 0# And horizon > size Then factor = horizon / size
    For attempt = 1 To 2
        If TryUltimateDirection(context, section, concreteMaterial, steelMaterial, loadPath, _
                eps0, kx, ky, rPath1, rPath2, rLimit, solver, lambdaValue, state, _
                correction(1) + factor * tangent(1), _
                (correction(2) + factor * tangent(2)) / mFrame.LengthScale, _
                (correction(3) + factor * tangent(3)) / mFrame.LengthScale) Then
            TryUltimateProjectedDirection = True
            Exit Function
        End If
        If context.StoppedTerminally Or factor = 1# Then Exit Function
        factor = 1#
    Next attempt
End Function

' Нормальная поправка достигает активного предела, два касательных
' направления сохраняют его. Масштабированные координаты имеют единицы деформации.
Private Function PrepareUltimateFace(ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal residual As Double, ByVal limitEps As Double, ByVal limitKx As Double, ByVal limitKy As Double, _
        ByRef q() As Double, ByRef correction() As Double, ByRef v() As Double) As Boolean
    Dim g(1 To 3) As Double, pivot As Long, i As Long, j As Long
    g(1) = limitEps: g(2) = limitKx / mFrame.LengthScale: g(3) = limitKy / mFrame.LengthScale
    pivot = 1
    For i = 2 To 3
        If Abs(g(i)) > Abs(g(pivot)) Then pivot = i
    Next i
    If g(pivot) = 0# Then Exit Function
    correction(pivot) = -residual / g(pivot)
    q(1) = eps0 + correction(1)
    q(2) = kx * mFrame.LengthScale + correction(2)
    q(3) = ky * mFrame.LengthScale + correction(3)
    For i = 1 To 3
        If i <> pivot Then
            j = j + 1
            v(i, j) = 1#: v(pivot, j) = -g(i) / g(pivot)
        End If
    Next i
    PrepareUltimateFace = True
End Function

' При недостаточном ранге локальной матрицы движение вдоль ее нулевого
' направления включает следующий участок материала, не меняя физический предел.
' Длина берется из реальных узлов диаграмм, а не из подобранной стартовой плоскости.
Private Function TryUltimateMaterialStep(ByVal context As ILimitSearchProblem, ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, ByVal loadPath As CLoadPathVector, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double, ByRef rPath1 As Double, _
        ByRef rPath2 As Double, ByRef rLimit As Double, ByRef solver As CSectionSolver, _
        ByRef lambdaValue As Double, ByRef state As String, ByVal h As Double, _
        ByVal limitEps As Double, ByVal limitKx As Double, ByVal limitKy As Double) As Boolean
    If Not TypeOf concreteMaterial Is CMaterialDiagram Or Not TypeOf steelMaterial Is CMaterialDiagram Then Exit Function
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram
    Set concrete = concreteMaterial: Set steel = steelMaterial
    Dim q(1 To 3) As Double, correction(1 To 3) As Double, v(1 To 3, 1 To 2) As Double
    If Not PrepareUltimateFace(eps0, kx, ky, rLimit, limitEps, limitKx, limitKy, q, correction, v) Then Exit Function
    Dim a(1 To 2, 1 To 2) As Double, j As Long
    For j = 1 To 2
        If Not BuildUltimateProjectedColumn(context, section, concreteMaterial, steelMaterial, loadPath, _
                q, v, j, h, a(1, j), a(2, j)) Then Exit Function
    Next j
    Dim ray As Long, sign As Long, i As Long, t1 As Double, t2 As Double, size As Double
    Dim direction(1 To 3) As Double, first As Double, second As Double, distance As Double
    Dim baseNorm As Double
    baseNorm = StepResidualNorm(context, rPath1, rPath2, rLimit)
    For ray = 1 To 4
        Select Case ray
            Case 1: t1 = 1#: t2 = 0#
            Case 2: t1 = 0#: t2 = 1#
            Case Else: t1 = -a(ray - 2, 2): t2 = a(ray - 2, 1)
        End Select
        For sign = 1 To -1 Step -2
            size = 0#
            For i = 1 To 3
                direction(i) = sign * (v(i, 1) * t1 + v(i, 2) * t2)
                size = MaxDouble(size, Abs(direction(i)))
            Next i
            If size > 0# Then
                For i = 1 To 3
                    direction(i) = direction(i) / size
                Next i
                If NonlinearMaterialTransitions(section, mFrame, concrete, steel, q, direction, first, second) Then
                    If second = 0# Then second = 2# * first
                    distance = first + (second - first) / 2#
                    If distance > first Then
                        If TryUltimatePlateauBracket(context, section, concreteMaterial, steelMaterial, loadPath, _
                                eps0, kx, ky, rPath1, rPath2, rLimit, solver, lambdaValue, state, _
                                correction(1) + distance * direction(1), _
                                (correction(2) + distance * direction(2)) / mFrame.LengthScale, _
                                (correction(3) + distance * direction(3)) / mFrame.LengthScale, _
                                first / distance, 1#, baseNorm) Then
                            TryUltimateMaterialStep = True
                            Exit Function
                        End If
                        If context.StoppedTerminally Then Exit Function
                    End If
                End If
            End If
        Next sign
    Next ray
End Function

' Двусторонняя проба вдоль касательного направления сохраняет активный предел.
Private Function BuildUltimateProjectedColumn(ByVal context As ILimitSearchProblem, ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, ByVal loadPath As CLoadPathVector, _
        ByRef q() As Double, ByRef v() As Double, ByVal column As Long, _
        ByVal h As Double, ByRef j1 As Double, ByRef j2 As Double) As Boolean
    Dim trial As CSectionSolver, p1 As Double, p2 As Double, p3 As Double
    Dim n1 As Double, n2 As Double, n3 As Double, trialLambda As Double, trialState As String
    If Not EvaluateResidual(context, section, concreteMaterial, steelMaterial, loadPath, _
            q(1) + h * v(1, column), (q(2) + h * v(2, column)) / mFrame.LengthScale, _
            (q(3) + h * v(3, column)) / mFrame.LengthScale, trial, p1, p2, p3, trialLambda, trialState) Then Exit Function
    If Not EvaluateResidual(context, section, concreteMaterial, steelMaterial, loadPath, _
            q(1) - h * v(1, column), (q(2) - h * v(2, column)) / mFrame.LengthScale, _
            (q(3) - h * v(3, column)) / mFrame.LengthScale, trial, n1, n2, n3, trialLambda, trialState) Then Exit Function
    j1 = (p1 - n1) / 2#: j2 = (p2 - n2) / 2#
    BuildUltimateProjectedColumn = True
End Function

' Проверяет обычную или регуляризованную поправку по настоящим невязкам.
Private Function TryUltimateDirection(ByVal context As ILimitSearchProblem, ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, ByVal loadPath As CLoadPathVector, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double, ByRef rPath1 As Double, _
        ByRef rPath2 As Double, ByRef rLimit As Double, ByRef solver As CSectionSolver, _
        ByRef lambdaValue As Double, ByRef state As String, ByVal dEps As Double, _
        ByVal dKx As Double, ByVal dKy As Double) As Boolean
    On Error GoTo Failed
    Dim baseNorm As Double
    baseNorm = StepResidualNorm(context, rPath1, rPath2, rLimit)

    Dim alpha As Double
    alpha = 1#
    Dim changingAlpha As Double

    Dim backtracks As Long
    Do While backtracks < NONLINEAR_MAX_BACKTRACKS
        backtracks = backtracks + 1
        Dim trialEps As Double, trialKx As Double, trialKy As Double
        trialEps = eps0 + alpha * dEps
        trialKx = kx + alpha * dKx
        trialKy = ky + alpha * dKy
        If trialEps = eps0 And trialKx = kx And trialKy = ky Then
            context.SetFailure sfcNumericalFailure, _
                "UltimateStrain остановлен: приращение плоскости деформаций меньше представимого шага Double, требуемые допуски не достигнуты."
            Exit Function
        End If
        Dim trialSolver As CSectionSolver
        Dim trialR1 As Double, trialR2 As Double, trialRLimit As Double
        Dim trialLambda As Double, trialState As String
        If EvaluateResidual(context, section, concreteMaterial, steelMaterial, _
                loadPath, trialEps, trialKx, trialKy, _
                trialSolver, trialR1, trialR2, trialRLimit, trialLambda, trialState) Then
            ' Уже достигнутые допуски важнее колебания нормы на уровне
            ' round-off. Итоговая приемка следующей итерации дополнительно
            ' подтверждает все усилия и физическую допустимость точки.
            If context.LimitSearchUltimateResidualReached(trialR1, trialR2, trialRLimit) Or _
                    NonlinearStepImproves(baseNorm, StepResidualNorm(context, trialR1, trialR2, trialRLimit), alpha) Then
                eps0 = trialEps
                kx = trialKx
                ky = trialKy
                Set solver = trialSolver
                rPath1 = trialR1: rPath2 = trialR2: rLimit = trialRLimit
                lambdaValue = trialLambda: state = trialState
                TryUltimateDirection = True
                Exit Function
            End If
            ' Между неудачной пробой вне плато и пробой на плато может
            ' лежать узкий допустимый шаг, который деление alpha пополам
            ' пропускает. Уточняем именно этот интервал по настоящим невязкам.
            If trialR1 = rPath1 And trialR2 = rPath2 Then
                If changingAlpha > alpha Then
                    If TryUltimatePlateauBracket(context, section, concreteMaterial, steelMaterial, loadPath, _
                            eps0, kx, ky, rPath1, rPath2, rLimit, solver, lambdaValue, state, _
                            dEps, dKx, dKy, alpha, changingAlpha, baseNorm) Then
                        TryUltimateDirection = True
                        Exit Function
                    End If
                    changingAlpha = 0#
                End If
            Else
                changingAlpha = alpha
            End If
        End If
        If context.StoppedTerminally Then Exit Function
        Dim nextAlpha As Double
        nextAlpha = alpha / 2#
        If nextAlpha <= 0# Or nextAlpha >= alpha Then
            context.SetFailure sfcNumericalFailure, "UltimateStrain остановлен: дальнейшее уменьшение шага демпфирования невозможно в машинной арифметике."
            Exit Function
        End If
        alpha = nextAlpha
    Loop

    Exit Function

Failed:
    ReportArithmeticOrContractFailure context, "Проверка поправки UltimateStrain", Err.Number, Err.Description
End Function

' Уточняет переход с плато на активную ветвь. Нулевая производная и
' неудачная крупная проба не означают, что между ними нет полезного шага.
Private Function TryUltimatePlateauBracket(ByVal context As ILimitSearchProblem, ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, ByVal loadPath As CLoadPathVector, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double, ByRef rPath1 As Double, _
        ByRef rPath2 As Double, ByRef rLimit As Double, ByRef solver As CSectionSolver, _
        ByRef lambdaValue As Double, ByRef state As String, ByVal dEps As Double, _
        ByVal dKx As Double, ByVal dKy As Double, ByVal lower As Double, ByVal upper As Double, _
        ByVal baseNorm As Double) As Boolean
    Dim attempt As Long, alpha As Double, trialE As Double, trialX As Double, trialY As Double
    Dim trial As CSectionSolver, r1 As Double, r2 As Double, r3 As Double
    Dim trialLambda As Double, trialState As String, trialNorm As Double
    For attempt = 1 To NONLINEAR_MAX_BACKTRACKS
        alpha = lower + (upper - lower) / 2#
        If alpha <= lower Or alpha >= upper Then Exit Function
        trialE = eps0 + alpha * dEps: trialX = kx + alpha * dKx: trialY = ky + alpha * dKy
        If Not EvaluateResidual(context, section, concreteMaterial, steelMaterial, loadPath, _
                trialE, trialX, trialY, trial, r1, r2, r3, trialLambda, trialState) Then Exit Function
        trialNorm = StepResidualNorm(context, r1, r2, r3)
        If context.LimitSearchUltimateResidualReached(r1, r2, r3) Or _
                NonlinearStepImproves(baseNorm, trialNorm, alpha) Then
            eps0 = trialE: kx = trialX: ky = trialY
            rPath1 = r1: rPath2 = r2: rLimit = r3
            Set solver = trial: lambdaValue = trialLambda: state = trialState
            TryUltimatePlateauBracket = True
            Exit Function
        End If
        If context.StoppedTerminally Then Exit Function
        If (r1 = rPath1 And r2 = rPath2) Or trialNorm < baseNorm Then
            lower = alpha
        Else
            upper = alpha
        End If
    Next attempt
End Function

' Строит один столбец численного Якобиана для общего UltimateStrain.
Private Function BuildUltimateJacobianColumn(ByVal context As ILimitSearchProblem, ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal eps0 As Double, _
        ByVal kx As Double, ByVal ky As Double, ByVal stepValue As Double, _
        ByVal rPath1 As Double, ByVal rPath2 As Double, ByVal rLimit As Double, _
        ByRef j1 As Double, ByRef j2 As Double, ByRef j3 As Double) As Boolean
    Dim solver As CSectionSolver
    Dim pR1 As Double, pR2 As Double, pLimit As Double
    Dim pLambda As Double, pState As String
    If Not EvaluateResidual(context, section, concreteMaterial, steelMaterial, _
            loadPath, eps0, kx, ky, solver, pR1, pR2, pLimit, pLambda, pState) Then Exit Function
    j1 = (pR1 - rPath1) / stepValue
    j2 = (pR2 - rPath2) / stepValue
    j3 = (pLimit - rLimit) / stepValue
    BuildUltimateJacobianColumn = True
End Function

' Переводит только представление плоскости. Усилия, траектория нагрузки,
' критерий и допуски сохраняются в исходной системе координат.
Private Function EvaluateResidual(ByVal context As ILimitSearchProblem, ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal centerStrain As Double, _
        ByVal kx As Double, ByVal ky As Double, ByRef solver As CSectionSolver, _
        ByRef rPath1 As Double, ByRef rPath2 As Double, ByRef rLimit As Double, _
        ByRef lambdaValue As Double, ByRef state As String) As Boolean
    EvaluateResidual = context.LimitSearchEvaluateUltimateResidual(section, concreteMaterial, steelMaterial, _
        loadPath, centerStrain, kx, ky, _
        solver, rPath1, rPath2, rLimit, lambdaValue, state, True)
End Function

' Переполнение или деление на ноль в численном шаге не означают физический
' предел. Другие VBA-ошибки относятся к программному контракту; исходные
' номер/текст сохраняются в журнале, а пользователь получает русскую причину.
Private Sub ReportArithmeticOrContractFailure(ByVal context As ILimitSearchProblem, _
        ByVal stage As String, ByVal errorNumber As Long, ByVal description As String)
    If context.StoppedTerminally Then Exit Sub
    Dim failureCode As ESolverFailureCode
    Dim reason As String
    context.AppendDiagnostic "ultimateRuntimeError=" & CStr(errorNumber) & "; stage=" & stage & _
        "; description=" & description
    If errorNumber = 6 Or errorNumber = 11 Then
        failureCode = sfcNumericalFailure
        reason = stage & ": не удалось вычислить численный шаг; число выходит за допустимый диапазон арифметики."
    Else
        failureCode = sfcInternalError
        reason = stage & ": внутренняя ошибка вычислений. Подробности сохранены в диагностическом журнале."
    End If
    ' Стандартное сообщение уже объяснено по-русски. Явно заданная причина
    ' callback-а содержит его контекст и сохраняется без изменения; текст
    ' используется только для представления, не для назначения статуса.
    Dim standardDescription As String
    ' При активном Err функция Error$(n) может вернуть custom Description.
    ' Номер и исходная причина уже сохранены аргументами и журналом.
    Err.Clear
    If errorNumber >= 0 And errorNumber <= 65535 Then standardDescription = Error$(errorNumber)
    If Len(description) > 0 And StrComp(description, standardDescription, vbBinaryCompare) <> 0 Then
        reason = reason & " Причина: " & description
    End If
    context.SetFailure failureCode, reason
End Sub

' Проверяет конечное число перед численным шагом; расчёт не зависит от Excel.
Private Function IsFiniteNumber(ByVal value As Double) As Boolean
    If value <> value Then Exit Function
    If Abs(value) > 1.797693134862315E+308 Then Exit Function
    IsFiniteNumber = True
End Function

' Для подбора шага невязки нормируются теми же строками, что и система Newton.
' Общий множитель ограничивает веса единицей; внутри line-search они неизменны.
' У вырожденного Якобиана сохраняется доменная норма. Прием точки отдельно
' проверяет исходные допуски усилий и деформаций, без этих весов.
Private Sub ConfigureStepResidualScales(ByVal scale1 As Double, ByVal scale2 As Double, ByVal scale3 As Double)
    mMeritScalingReady = (scale1 > 0# And scale2 > 0# And scale3 > 0#)
    If Not mMeritScalingReady Then Exit Sub
    Dim minimum As Double
    minimum = MinDouble(scale1, MinDouble(scale2, scale3))
    mMeritWeight1 = minimum / scale1
    mMeritWeight2 = minimum / scale2
    mMeritWeight3 = minimum / scale3
    mMeritScalingReady = (mMeritWeight1 > 0# And mMeritWeight2 > 0# And mMeritWeight3 > 0#)
End Sub

Private Function StepResidualNorm(ByVal context As ILimitSearchProblem, _
        ByVal r1 As Double, ByVal r2 As Double, ByVal r3 As Double) As Double
    If mMeritScalingReady Then
        StepResidualNorm = Abs(r1) * mMeritWeight1 + Abs(r2) * mMeritWeight2 + Abs(r3) * mMeritWeight3
    Else
        StepResidualNorm = context.LimitSearchUltimateResidualNorm(r1, r2, r3)
    End If
End Function

' Утилитарный минимум без обращения к Excel.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

' Утилитарный максимум без обращения к Excel.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function
