VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLoadMultiplierSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Общий LoadMultiplier-поиск предельной lambda-точки
' ==========================================================================
' Класс владеет низкоуровневой одномерной search-логикой:
' расширением физической скобки, Bisection/Brent/Secant, защитой Secant-кандидата
' и общей финализацией выбранной lambda со стороны acceptable-ветки.
'
' Capacity и CrackFormation отличаются не алгоритмом поиска, а только тем, как
' оценивается одна lambda-точка и что означает найденный предел. Поэтому класс
' спрашивает у доменного объекта значение функции и классификацию точки, а сам
' управляет порядком пробных lambda.

' Выполняет общий LoadMultiplier независимо от инженерного вида задачи.
' Подготовка, начальный физический исход и приемка точки принадлежат callback-у.
Public Function Execute(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request Is Nothing Then Err.Raise vbObjectError + 4220, "CLoadMultiplierSearch", "Не передан запрос limit-search."
    If request.Problem Is Nothing Then
        Set Execute = ContractFailure(request, "Не передана задача общего поиска.")
        Exit Function
    End If
    If Not request.Problem.Prepare(request) Then
        Set Execute = request.Problem.BuildResult(request, False)
        Exit Function
    End If

    Dim loadPath As CLoadPathVector
    Set loadPath = request.Problem.BuildLoadPath(request)
    If loadPath Is Nothing Then
        Set Execute = ContractFailure(request, "Задача поиска не передала траекторию нагрузки.")
        Exit Function
    End If
    RunSearch request, loadPath
    Set Execute = request.Problem.BuildResult(request, request.Problem.Converged)
End Function

' Отсутствие обязательного внутреннего объекта является ошибкой контракта,
' а не ошибкой ввода пользователя и не численной несходимостью.
Private Function ContractFailure(ByVal request As CLimitSearchRequest, ByVal reason As String) As CLimitSearchResult
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetResult rsInternalError, rcInternalError, request.ResultKind, reason, vbNullString, True, False
    Dim result As CLimitSearchResult
    Set result = New CLimitSearchResult
    result.Initialize meta, request.SolutionStrategy, vbNullString, False, 0#, 0#, 0#, 0#, _
        Nothing, vbNullString, vbNullString, vbNullString, False
    Set ContractFailure = result
End Function

' Решает начальную точку, находит скобку и запускает выбранный численный метод.
' Здесь нет проверки вида Capacity/Crack, нагрузки текущего LC или psi.
Private Sub RunSearch(ByVal request As CLimitSearchRequest, ByVal loadPath As CLoadPathVector)
    If Not ValidateSearchSettings(request) Then Exit Sub
    Dim lowLambda As Double, highLambda As Double
    Dim lowSolver As CSectionSolver, highSolver As CSectionSolver
    Dim functionValue As Double, probeState As String
    lowLambda = 0#
    highLambda = request.Problem.InitialLambdaStep

    If Not request.Problem.EvaluatePoint(request, loadPath, Nothing, lowLambda, _
            functionValue, probeState, lowSolver) Then
        request.Problem.HandleInitialFailure probeState
        Exit Sub
    End If
    If Not request.Problem.AcceptInitialPoint(request, loadPath, functionValue, probeState, lowSolver) Then Exit Sub

    If Not FindBracket(request, loadPath, lowLambda, lowSolver, highLambda, highSolver) Then
        request.Problem.TryFinalizeAfterSearchFailure request
        Exit Sub
    End If
    RunSearchInsideBracket request, loadPath, lowLambda, highLambda, lowSolver, highSolver, _
        request.Problem.SearchMethod, request.Problem.LambdaTolerance, request.Problem.MaxIterations
    If Not request.Problem.Converged Then request.Problem.TryFinalizeAfterSearchFailure request
End Sub

' Проверяет численный диапазон до проб и арифметики. Нулевой bisection-budget
' означает уточнение до заданной точности с обязательным контролем прогресса.
Private Function ValidateSearchSettings(ByVal request As CLimitSearchRequest) As Boolean
    If Not IsFiniteNumber(request.Problem.InitialLambdaStep) Or _
            Not IsFiniteNumber(request.Problem.MaxLambda) Or _
            Not IsFiniteNumber(request.Problem.LambdaTolerance) Then
        request.Problem.SetInvalidInput "Границы и допуск поиска lambda должны быть конечными числами."
    ElseIf request.Problem.InitialLambdaStep <= 0# Or request.Problem.MaxLambda <= 0# Or _
            request.Problem.LambdaTolerance <= 0# Then
        request.Problem.SetInvalidInput "Начальный шаг, MaxLambda и допуск lambda должны быть положительными."
    Else
        ValidateSearchSettings = True
    End If
End Function

' Единое расширение скобки: допустимая проба сдвигает нижнюю границу,
' достигнутый критерий задает верхнюю, а численная неудача требует recovery.
' Последняя проба всегда выполняется на MaxLambda, без пропуска хвоста интервала.
Private Function FindBracket(ByVal request As CLimitSearchRequest, ByVal loadPath As CLoadPathVector, _
        ByRef lowLambda As Double, ByRef lowSolver As CSectionSolver, _
        ByRef highLambda As Double, ByRef highSolver As CSectionSolver) As Boolean
    On Error GoTo ArithmeticFailed
    Dim functionValue As Double, probeState As String
    Dim evaluated As Boolean, reducedLambda As Double
    highLambda = MinDouble(highLambda, request.Problem.MaxLambda)
    Do
        evaluated = request.Problem.EvaluatePoint(request, loadPath, lowSolver, highLambda, _
            functionValue, probeState, highSolver)
        If Not evaluated Or request.Problem.IsNumericalFailureState(probeState) Then
            If request.Problem.StoppedTerminally Then Exit Function
            reducedLambda = request.Problem.ReduceLambdaAfterNumericalFailure(lowLambda, highLambda, probeState)
            If Not IsFiniteNumber(reducedLambda) Or reducedLambda <= lowLambda Or reducedLambda >= highLambda Then
                request.Problem.SetNumericalFailure _
                    "Восстановление поиска lambda не уменьшило интервал; следующую пробу невозможно выбрать."
                Exit Function
            End If
            If reducedLambda - lowLambda <= request.Problem.LambdaTolerance Then
                request.Problem.HandleBracketNumericalFailure lowLambda, highLambda, probeState
                Exit Function
            End If
            highLambda = reducedLambda
        ElseIf request.Problem.IsLimitState(probeState) Then
            FindBracket = True
            Exit Function
        ElseIf request.Problem.IsAcceptableState(probeState) Then
            lowLambda = highLambda
            Set lowSolver = highSolver
            request.Problem.RecordAcceptable loadPath, lowLambda
            If highLambda >= request.Problem.MaxLambda Then Exit Do
            ' Успешная проба дает более близкий warm-start. Несошедшаяся
            ' нагрузка не является физической границей: повторяем расширение.
            If highLambda >= request.Problem.MaxLambda / 2# Then
                highLambda = request.Problem.MaxLambda
            Else
                highLambda = 2# * highLambda
            End If
            If highLambda <= lowLambda Then
                request.Problem.SetNumericalFailure "Расширение интервала lambda остановилось из-за машинной точности."
                Exit Function
            End If
        Else
            request.Problem.HandleUnexpectedState probeState
            Exit Function
        End If
    Loop
    request.Problem.HandleBracketNotFound lowLambda
    Exit Function
ArithmeticFailed:
    request.Problem.SetNumericalFailure "Не удалось продолжить численный поиск lambda: " & Err.Description
End Function

' Запускает выбранный одномерный метод внутри уже найденной физической скобки.
Private Sub RunSearchInsideBracket(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lowLambda As Double, _
        ByVal highLambda As Double, ByVal lowSolver As CSectionSolver, _
        ByVal highSolver As CSectionSolver, ByVal methodName As String, _
        ByVal tolerance As Double, ByVal maxIterations As Long)
    Select Case methodName
        Case "Bisection"
            Dim iterations As Long
            request.Problem.AppendDiagnostic "searchMethod=Bisection"
            If SearchByBisectionCore(request, loadPath, lowLambda, lowSolver, highLambda, highSolver, _
                    tolerance, request.Problem.BisectionMaxIterations, iterations) Then
                request.Problem.FinalizeBracket request, loadPath, lowLambda, lowSolver, _
                    highLambda, highSolver, "Bisection", iterations
            End If
        Case "Brent"
            SearchByBrent request, loadPath, lowLambda, highLambda, tolerance, maxIterations
        Case "Secant"
            SearchBySecant request, loadPath, lowLambda, highLambda, tolerance, maxIterations
        Case Else
            request.Problem.SetInvalidInput _
                "Метод одномерного поиска должен быть Bisection, Brent или Secant."
    End Select
End Sub

' Выполняет общий цикл bisection для уже найденной lambda-скобки.
' Capacity и CrackFormation используют один порядок пробных lambda-точек; различаются только
' callback оценки точки, смысл статуса предела и то, какая граница затем финализируется.
Private Function SearchByBisectionCore(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByRef lowLambda As Double, _
        ByRef lowSolver As CSectionSolver, ByRef highLambda As Double, _
        ByRef highSolver As CSectionSolver, ByVal tolerance As Double, _
        ByVal maxIterations As Long, ByRef iterations As Long) As Boolean
    iterations = 0
    Do While (highLambda - lowLambda) > tolerance
        If maxIterations > 0 And iterations >= maxIterations Then Exit Do

        Dim midLambda As Double
        Dim midSolver As CSectionSolver
        Dim midResidual As Double
        Dim probeState As String
        Dim startSolver As CSectionSolver

        midLambda = SafeMidpoint(lowLambda, highLambda)
        If midLambda <= lowLambda Or midLambda >= highLambda Then
            request.Problem.SetNumericalFailure _
                "Поиск Bisection остановлен: между границами lambda нет представимого числа, а заданная точность еще не достигнута."
            Exit Function
        End If
        Set startSolver = PreferredBisectionStartSolver(lowSolver, highSolver)
        If Not EvaluateBisectionPoint(request, loadPath, startSolver, midLambda, _
                midResidual, probeState, midSolver) Then
            HandleBisectionEvaluationFailure request, midLambda
            Exit Function
        End If
        iterations = iterations + 1

        If IsAcceptableState(request, probeState) Then
            lowLambda = midLambda
            Set lowSolver = midSolver
        ElseIf IsLimitState(request, probeState) Then
            highLambda = midLambda
            Set highSolver = midSolver
        ElseIf IsNumericalFailureState(request, probeState) Then
            HandleBisectionNumericalFailure request, lowLambda, midLambda, probeState
            Exit Function
        Else
            HandleBisectionUnexpectedState request, probeState
            Exit Function
        End If
    Loop

    If (highLambda - lowLambda) > tolerance Then
        HandleBisectionIterationLimit request, maxIterations, highLambda - lowLambda, tolerance
        Exit Function
    End If

    SearchByBisectionCore = True
End Function

' Защищенный Brent: пробует интерполяционный шаг и возвращается к скобке,
' если интерполяция становится рискованной.
Private Sub SearchByBrent(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lowLambda As Double, _
        ByVal highLambda As Double, ByVal tolerance As Double, ByVal maxIterations As Long)
    request.Problem.AppendDiagnostic "searchMethod=Brent"

    Dim a As Double, b As Double, c As Double
    Dim fa As Double, fb As Double, fc As Double
    Dim stateA As String, stateB As String, stateC As String
    Dim solverA As CSectionSolver, solverB As CSectionSolver, solverC As CSectionSolver
    Dim d As Double, e As Double
    Dim iter As Long
    a = lowLambda
    b = highLambda
    If Not EvaluateLimitPoint(request, loadPath, a, fa, stateA, solverA) Then Exit Sub
    If Not EvaluateLimitPoint(request, loadPath, b, fb, stateB, solverB) Then Exit Sub
    If fb = 0# And IsLimitState(request, stateB) Then
        If FinalizeKnownProbe(request, loadPath, b, stateB, "Brent", solverB, 0) Then Exit Sub
        If request.Problem.StoppedTerminally Then Exit Sub
    End If
    If fa >= 0# Or fb <= 0# Then
        request.Problem.SetInvalidInput _
            "Методу Brent нужна скобка: нижняя lambda с запасом и верхняя lambda за физическим пределом."
        Exit Sub
    End If

    c = a
    fc = fa
    stateC = stateA
    Set solverC = solverA
    d = b - a
    e = d

    For iter = 1 To maxIterations
        If Sgn(fb) = Sgn(fc) Then
            c = a
            fc = fa
            stateC = stateA
            Set solverC = solverA
            d = b - a
            e = d
        End If
        If Abs(fc) < Abs(fb) Then
            a = b: b = c: c = a
            fa = fb: fb = fc: fc = fa
            stateA = stateB: stateB = stateC: stateC = stateA
            Set solverA = solverB: Set solverB = solverC: Set solverC = solverA
        End If

        Dim tol As Double
        Dim xm As Double
        tol = tolerance / 2#
        xm = (c - b) / 2#
        If fb = 0# And IsLimitState(request, stateB) Then
            If FinalizeKnownProbe(request, loadPath, b, stateB, "Brent", solverB, iter) Then Exit Sub
            If request.Problem.StoppedTerminally Then Exit Sub
        End If
        If Abs(c - b) <= tolerance Then
            ' Brent выбирает ближайшую к нулю функцию, а не обязательно
            ' допустимую сторону. Передаем обе сохраненные стороны доменному
            ' callback: Capacity принимает нижнюю, Formation подтверждает верхнюю.
            If IsAcceptableState(request, stateB) Then
                request.Problem.FinalizeBracket request, loadPath, b, solverB, c, solverC, "Brent", iter
            ElseIf IsAcceptableState(request, stateC) Then
                request.Problem.FinalizeBracket request, loadPath, c, solverC, b, solverB, "Brent", iter
            Else
                request.Problem.SetNumericalFailure "В конечной скобке Brent не сохранена допустимая точка равновесия."
            End If
            Exit Sub
        End If

        If Abs(e) >= tol And Abs(fa) > Abs(fb) Then
            Dim interpolatedStep As Double
            If TryBrentStep(a, b, c, fa, fb, fc, xm, tol, e, interpolatedStep) Then
                e = d
                d = interpolatedStep
            Else
                d = xm
                e = d
            End If
        Else
            d = xm
            e = d
        End If

        a = b
        fa = fb
        stateA = stateB
        Set solverA = solverB
        Dim candidate As Double
        If Not TryBrentCandidate(b, c, d, tol, xm, candidate) Then
            request.Problem.SetNumericalFailure _
                "Поиск Brent остановлен: между границами lambda нет представимого шага заданной точности."
            Exit Sub
        End If
        b = candidate
        If Not EvaluateLimitPoint(request, loadPath, b, fb, stateB, solverB) Then Exit Sub
    Next iter

    request.Problem.SetNumericalFailure "Поиск Brent достиг лимита итераций."
End Sub

' Вычисляет штатную интерполяцию Brent только при безопасной арифметике.
' Переполнение или вырождение отношения отвергает интерполяцию, а не скобку:
' вызывающий алгоритм продолжает обычным бисекционным шагом.
Private Function TryBrentStep(ByVal a As Double, ByVal b As Double, ByVal c As Double, _
        ByVal fa As Double, ByVal fb As Double, ByVal fc As Double, ByVal xm As Double, _
        ByVal tol As Double, ByVal previousStep As Double, ByRef result As Double) As Boolean
    On Error GoTo UnsafeInterpolation
    Dim s As Double, p As Double, q As Double, r As Double
    If fa = 0# Or fc = 0# Then Exit Function
    s = fb / fa
    If a = c Then
        p = 2# * xm * s
        q = 1# - s
    Else
        q = fa / fc
        r = fb / fc
        p = s * (2# * xm * q * (q - r) - (b - a) * (r - 1#))
        q = (q - 1#) * (r - 1#) * (s - 1#)
    End If
    If q = 0# Then Exit Function
    If p > 0# Then q = -q
    p = Abs(p)
    If 2# * p >= MinDouble(3# * xm * q - Abs(tol * q), Abs(previousStep * q)) Then Exit Function
    result = p / q
    TryBrentStep = IsFiniteNumber(result)
UnsafeInterpolation:
End Function

' Удерживает следующий шаг Brent внутри текущей, а не исходной скобки.
' Если сложение или минимальный шаг не представимы, пробует безопасную середину.
Private Function TryBrentCandidate(ByVal current As Double, ByVal opposite As Double, _
        ByVal delta As Double, ByVal tol As Double, ByVal direction As Double, _
        ByRef candidate As Double) As Boolean
    On Error GoTo UseMidpoint
    If Abs(delta) > tol Then
        candidate = current + delta
    Else
        candidate = current + Sgn(direction) * tol
    End If
    If IsFiniteNumber(candidate) Then
        If candidate > MinDouble(current, opposite) And candidate < MaxDouble(current, opposite) Then
            TryBrentCandidate = True
            Exit Function
        End If
    End If
UseMidpoint:
    candidate = SafeMidpoint(current, opposite)
    TryBrentCandidate = candidate > MinDouble(current, opposite) And candidate < MaxDouble(current, opposite)
End Function

' Метод секущих: строит следующую lambda по двум предыдущим значениям функции,
' но дополнительно удерживает точку внутри физической скобки.
Private Sub SearchBySecant(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lowLambda As Double, _
        ByVal highLambda As Double, ByVal tolerance As Double, ByVal maxIterations As Long)
    request.Problem.AppendDiagnostic "searchMethod=Secant"

    Dim x0 As Double, x1 As Double, x2 As Double
    Dim f0 As Double, f1 As Double, f2 As Double
    Dim state0 As String, state1 As String, state2 As String
    Dim solver0 As CSectionSolver, solver1 As CSectionSolver, solver2 As CSectionSolver
    Dim bracketLow As Double
    Dim bracketHigh As Double
    Dim bracketLowSolver As CSectionSolver
    Dim bracketHighSolver As CSectionSolver
    Dim iter As Long

    x0 = lowLambda
    x1 = highLambda
    If Not EvaluateLimitPoint(request, loadPath, x0, f0, state0, solver0) Then Exit Sub
    If Not EvaluateLimitPoint(request, loadPath, x1, f1, state1, solver1) Then Exit Sub
    If Abs(f1) <= tolerance And CanFinalizeSecantProbe(request, state1, solver1) Then
        If FinalizeKnownProbe(request, loadPath, x1, state1, "Secant", solver1, 0) Then Exit Sub
        If request.Problem.StoppedTerminally Then Exit Sub
    End If
    If f0 >= 0# Or f1 <= 0# Then
        request.Problem.SetInvalidInput _
            "Методу Secant нужна скобка: нижняя lambda с запасом и верхняя lambda за физическим пределом."
        Exit Sub
    End If
    bracketLow = x0
    bracketHigh = x1
    Set bracketLowSolver = solver0
    Set bracketHighSolver = solver1

    For iter = 1 To maxIterations
        If Abs(f1) <= tolerance And CanFinalizeSecantProbe(request, state1, solver1) Then
            If FinalizeKnownProbe(request, loadPath, x1, state1, "Secant", solver1, iter) Then Exit Sub
            If request.Problem.StoppedTerminally Then Exit Sub
        End If
        If Abs(bracketHigh - bracketLow) <= tolerance Then
            request.Problem.FinalizeBracket request, loadPath, bracketLow, bracketLowSolver, _
                bracketHigh, bracketHighSolver, "Secant", iter
            Exit Sub
        End If
        If Not TrySecantCandidate(x0, x1, f0, f1, x2) Then x2 = SafeMidpoint(bracketLow, bracketHigh)
        x2 = GuardSecantCandidate(request, x2, x1, bracketLow, bracketHigh)
        If x2 = x1 Then x2 = SafeMidpoint(bracketLow, bracketHigh)
        If x2 <= bracketLow Or x2 >= bracketHigh Or Not IsFiniteNumber(x2) Then
            request.Problem.SetNumericalFailure _
                "Поиск Secant остановлен: между границами lambda нет представимого шага заданной точности."
            Exit Sub
        End If

        If Not EvaluateLimitPoint(request, loadPath, x2, f2, state2, solver2) Then Exit Sub
        UpdateSecantBracket request, x2, f2, state2, solver2, _
            bracketLow, bracketHigh, bracketLowSolver, bracketHighSolver
        If Abs(f2) <= tolerance And CanFinalizeSecantProbe(request, state2, solver2) Then
            If FinalizeKnownProbe(request, loadPath, x2, state2, "Secant", solver2, iter) Then Exit Sub
            If request.Problem.StoppedTerminally Then Exit Sub
        End If
        If Abs(bracketHigh - bracketLow) <= tolerance Then
            request.Problem.FinalizeBracket request, loadPath, bracketLow, bracketLowSolver, _
                bracketHigh, bracketHighSolver, "Secant", iter
            Exit Sub
        End If

        x0 = x1
        f0 = f1
        state0 = state1
        Set solver0 = solver1
        x1 = x2
        f1 = f2
        state1 = state2
        Set solver1 = solver2
    Next iter

    request.Problem.SetNumericalFailure "Поиск Secant достиг лимита итераций."
End Sub

' Нормирует значения функции перед разностью, чтобы сумма больших величин
' противоположных знаков не переполнилась. Не меняет критерий точности;
' небезопасная интерполяция заменяется шагом внутри подтвержденной скобки.
Private Function TrySecantCandidate(ByVal x0 As Double, ByVal x1 As Double, _
        ByVal f0 As Double, ByVal f1 As Double, ByRef candidate As Double) As Boolean
    On Error GoTo UnsafeInterpolation
    Dim functionScale As Double, denominator As Double
    functionScale = MaxDouble(Abs(f0), Abs(f1))
    If functionScale = 0# Then Exit Function
    denominator = f1 / functionScale - f0 / functionScale
    If denominator = 0# Then Exit Function
    candidate = x1 - (f1 / functionScale) * (x1 - x0) / denominator
    TrySecantCandidate = IsFiniteNumber(candidate)
UnsafeInterpolation:
End Function

' Фиксирует исчерпание итераций bisection как численную неудачу, если скобка
' все еще шире заданного допуска. Иначе search мог бы выдать недоточненную
' нижнюю границу как успешную найденную точку.
Private Sub HandleBisectionIterationLimit(ByVal request As CLimitSearchRequest, _
        ByVal maxIterations As Long, ByVal remainingWidth As Double, ByVal tolerance As Double)
    Dim reason As String
    reason = "Поиск Bisection достиг лимита итераций до заданной точности: ширина скобки " & _
        FormatNumberInvariant(remainingWidth) & ", допуск " & FormatNumberInvariant(tolerance) & "."
    If maxIterations <= 0 Then _
        reason = "Поиск Bisection остановлен до достижения заданной точности: ширина скобки " & _
            FormatNumberInvariant(remainingWidth) & ", допуск " & FormatNumberInvariant(tolerance) & "."

    request.Problem.HandleIterationLimit maxIterations, remainingWidth, tolerance
End Sub

' Пересчитывает итоговую lambda и передает доменному объекту финализацию.
Private Function FinalizeAfterProbe(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal methodName As String) As Boolean
    Dim functionValue As Double
    Dim probeState As String
    Dim solver As CSectionSolver
    If Not request.Problem.EvaluatePoint(request, loadPath, Nothing, _
            lambdaValue, functionValue, probeState, solver) Then Exit Function
    FinalizeAfterProbe = request.Problem.FinalizeAt(request, loadPath, lambdaValue, probeState, methodName, solver, 0)
End Function

' Финализирует уже рассчитанную probe-точку без повторного solve.
' Это важно для Secant: верхняя точка физической скобки может быть найденной
' несошедшейся probe, пригодной как граница, но не как итоговое НДС. False
' означает отказ доменного callback-а; общий цикл продолжает уточнение.
Private Function FinalizeKnownProbe(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal probeState As String, ByVal methodName As String, _
        ByVal solver As CSectionSolver, ByVal iterations As Long) As Boolean
    If solver Is Nothing Then
        FinalizeKnownProbe = FinalizeAfterProbe(request, loadPath, lambdaValue, methodName)
    Else
        FinalizeKnownProbe = request.Problem.FinalizeAt(request, loadPath, lambdaValue, probeState, methodName, solver, iterations)
    End If
End Function

' Выбирает стартовое НДС для следующей bisection-пробы.
' Предпочтителен сошедшийся снимок с допустимой стороны; если его нет,
' используется подтвержденная верхняя точка. Принятие warm-start решает callback.
Private Function PreferredBisectionStartSolver(ByVal lowSolver As CSectionSolver, _
        ByVal highSolver As CSectionSolver) As CSectionSolver
    If lowSolver Is Nothing Then
        Set PreferredBisectionStartSolver = highSolver
    Else
        Set PreferredBisectionStartSolver = lowSolver
    End If
End Function

' Оценивает одну lambda-точку для общего bisection-цикла.
' Callback возвращает значение критерия и связанный с ним solver для
' warm-start и подтверждения конечной точки без повторного solve.
Private Function EvaluateBisectionPoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByVal lambdaValue As Double, ByRef functionValue As Double, _
        ByRef probeState As String, ByRef solver As CSectionSolver) As Boolean
    EvaluateBisectionPoint = request.Problem.EvaluatePoint(request, loadPath, _
        startSolver, lambdaValue, functionValue, probeState, solver)
End Function

' Возвращает доменному объекту понятную причину отказа пробной lambda-точки.
' Первичная окончательная причина callback-а должна сохраняться; численная
' ошибка относится только к пробе без подтвержденного равновесия.
Private Sub HandleBisectionEvaluationFailure(ByVal request As CLimitSearchRequest, _
        ByVal lambdaValue As Double)
    If request.Problem.StoppedTerminally Then Exit Sub
    request.Problem.SetNumericalFailure _
        "Не удалось найти равновесие при lambda = " & _
        FormatNumberInvariant(lambdaValue) & " в поиске предельной точки."
End Sub

' Обрабатывает статус NumericalFailure, если callback смог классифицировать точку без обычного False.
Private Sub HandleBisectionNumericalFailure(ByVal request As CLimitSearchRequest, _
        ByVal lowLambda As Double, ByVal failedLambda As Double, ByVal probeState As String)
    request.Problem.HandleBracketNumericalFailure lowLambda, failedLambda, probeState
End Sub

' Защищает общий bisection от неизвестного статуса пробной точки.
' Такая ситуация означает не физический результат, а нарушение контракта callback-а.
Private Sub HandleBisectionUnexpectedState(ByVal request As CLimitSearchRequest, _
        ByVal probeState As String)
    request.Problem.HandleUnexpectedState probeState
End Sub

' Оценивает lambda-точку и возвращает связанный с ней solver-снимок.
' Brent/Secant хранят этот снимок вместе с границей скобки, чтобы финализация не
' брала LastSolver от другой probe-точки.
Private Function EvaluateLimitPoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByRef functionValue As Double, ByRef probeState As String, _
        ByRef solver As CSectionSolver) As Boolean
    EvaluateLimitPoint = request.Problem.EvaluatePoint(request, loadPath, Nothing, _
        lambdaValue, functionValue, probeState, solver)
End Function

' Проверяет, что lambda-точка находится ниже предела.
Private Function IsAcceptableState(ByVal request As CLimitSearchRequest, ByVal probeState As String) As Boolean
    IsAcceptableState = request.Problem.IsAcceptableState(probeState)
End Function

' Проверяет, что lambda-точка находится на физическом пределе.
Private Function IsLimitState(ByVal request As CLimitSearchRequest, ByVal probeState As String) As Boolean
    IsLimitState = request.Problem.IsLimitState(probeState)
End Function

' Проверяет, что отказ lambda-точки является численной неудачей.
Private Function IsNumericalFailureState(ByVal request As CLimitSearchRequest, ByVal probeState As String) As Boolean
    IsNumericalFailureState = request.Problem.IsNumericalFailureState(probeState)
End Function

' Проверяет, можно ли использовать Secant-probe как итоговое НДС.
' Acceptable-точка уже находится с безопасной стороны. Limit-точка допустима
' только если внутренний CSectionSolver действительно сошелся; несошедшаяся
' физическая probe остается только верхней границей поиска.
Private Function CanFinalizeSecantProbe(ByVal request As CLimitSearchRequest, _
        ByVal probeState As String, ByVal solver As CSectionSolver) As Boolean
    If IsAcceptableState(request, probeState) Then
        CanFinalizeSecantProbe = True
    ElseIf IsLimitState(request, probeState) Then
        If solver Is Nothing Then
            ' Чисто аналитическая задача может подтвердить критерий без State.
            ' Реальное НДС и физическую допустимость проверяет ее FinalizeAt.
            CanFinalizeSecantProbe = True
        Else
            CanFinalizeSecantProbe = solver.Converged
        End If
    End If
End Function

' Удерживает Secant-кандидат внутри найденной физической скобки.
Private Function GuardSecantCandidate(ByVal request As CLimitSearchRequest, _
        ByVal rawCandidate As Double, ByVal currentLambda As Double, _
        ByVal bracketLow As Double, ByVal bracketHigh As Double) As Double
    GuardSecantCandidate = rawCandidate
    If rawCandidate > bracketLow And rawCandidate < bracketHigh Then Exit Function

    If rawCandidate <= bracketLow Then
        If currentLambda > bracketLow And currentLambda < bracketHigh Then
            GuardSecantCandidate = SafeMidpoint(bracketLow, currentLambda)
        Else
            GuardSecantCandidate = bracketLow + 0.25 * (bracketHigh - bracketLow)
        End If
    ElseIf rawCandidate >= bracketHigh Then
        If currentLambda > bracketLow And currentLambda < bracketHigh Then
            GuardSecantCandidate = SafeMidpoint(currentLambda, bracketHigh)
        Else
            GuardSecantCandidate = bracketLow + 0.75 * (bracketHigh - bracketLow)
        End If
    End If

    request.Problem.AppendDiagnostic _
        "secantGuard=insideBracket; rawLambda=" & FormatNumberInvariant(rawCandidate) & _
        "; guardedLambda=" & FormatNumberInvariant(GuardSecantCandidate) & _
        "; bracketLow=" & FormatNumberInvariant(bracketLow) & _
        "; bracketHigh=" & FormatNumberInvariant(bracketHigh)
End Function

' Обновляет защитную физическую скобку Secant.
Private Sub UpdateSecantBracket(ByVal request As CLimitSearchRequest, _
        ByVal lambdaValue As Double, ByVal functionValue As Double, _
        ByVal probeState As String, ByVal solver As CSectionSolver, _
        ByRef bracketLow As Double, ByRef bracketHigh As Double, _
        ByRef bracketLowSolver As CSectionSolver, ByRef bracketHighSolver As CSectionSolver)
    If functionValue < 0# Then
        If lambdaValue > bracketLow Then
            bracketLow = lambdaValue
            Set bracketLowSolver = solver
        End If
    Else
        If lambdaValue < bracketHigh Then
            bracketHigh = lambdaValue
            Set bracketHighSolver = solver
        End If
    End If
End Sub

' Проверяет, что число конечно и пригодно для дальнейшей арифметики.
Private Function IsFiniteNumber(ByVal value As Double) As Boolean
    If value <> value Then Exit Function
    If Abs(value) > 1.797693134862315E+308 Then Exit Function
    IsFiniteNumber = True
End Function

' Середина неотрицательных lambda-границ без переполнения их суммы.
' Порядок аргументов может быть обратным в локальной скобке Brent.
Private Function SafeMidpoint(ByVal first As Double, ByVal second As Double) As Double
    SafeMidpoint = first + (second - first) / 2#
End Function

' Форматирует число для диагностического журнала с точкой как разделителем.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    If value <> 0# And (Abs(value) < 0.000000001 Or Abs(value) > 1000000000000#) Then
        FormatNumberInvariant = Replace$(Format$(value, "0.############E+00"), ",", ".")
    Else
        FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
    End If
End Function

' Утилитарный минимум без WorksheetFunction в расчетном ядре.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

' Утилитарный максимум без WorksheetFunction в расчетном ядре.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function
