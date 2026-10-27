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
' Поиск предельной нагрузки изменением множителя lambda
' ==========================================================================
' Строит интервал между допустимой и предельной нагрузкой, затем уточняет
' его методами Bisection, Brent или Secant. Проверяет продвижение каждой пробы
' и передаёт выбранную точку задаче для окончательного подтверждения.
' Несущая способность и образование трещины используют общий алгоритм;
' физический критерий и оценку одной нагрузки предоставляет сама задача.

' Выполняет общий поиск предельной нагрузки методом LoadMultiplier.
' Инженерная задача задаёт подготовку, физический смысл начальной точки
' и окончательное подтверждение найденного предела.
Public Function Execute(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request Is Nothing Then Err.Raise vbObjectError + 4220, "CLoadMultiplierSearch", "Не переданы исходные данные для поиска предельного состояния."
    If request.Problem Is Nothing Then
        Set Execute = ContractFailure(request, "Не передана задача общего поиска.")
        Exit Function
    End If
    On Error GoTo SearchFailed
    If Not request.Problem.Prepare(request) Then
        On Error GoTo 0
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
    On Error GoTo 0
    Set Execute = request.Problem.BuildResult(request, request.Problem.Converged)
    Exit Function
SearchFailed:
    ReportArithmeticOrContractFailure request.Problem, _
        "Поиск предельного множителя нагрузки", Err.Number, Err.Description
    Set Execute = request.Problem.BuildResult(request, False)
End Function

' Отсутствие обязательного внутреннего объекта — ошибка программы,
' а не ошибка пользовательского ввода и не численная несходимость.
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

' Оценивает начальную нагрузку, находит допустимую и предельную стороны интервала
' и уточняет lambda выбранным методом. Критерий задаёт инженерная задача.
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
        If Not request.Problem.Converged And Not request.Problem.StoppedTerminally Then _
            request.Problem.TryFinalizeAfterSearchFailure request
        Exit Sub
    End If
    Dim pointAccepted As Boolean
    RunSearchInsideBracket request, loadPath, lowLambda, highLambda, lowSolver, highSolver, _
        request.Problem.SearchMethod, request.Problem.LambdaTolerance, request.Problem.MaxIterations, pointAccepted
    If Not pointAccepted And Not request.Problem.Converged Then
        If Not request.Problem.StoppedTerminally Then request.Problem.TryFinalizeAfterSearchFailure request
    End If
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
    Dim evaluated As Boolean, reducedLambda As Double, recoveryAttempts As Long
    Dim boundaryRefinementAttempted As Boolean
    highLambda = MinDouble(highLambda, request.Problem.MaxLambda)
    Do
        evaluated = request.Problem.EvaluatePoint(request, loadPath, lowSolver, highLambda, _
            functionValue, probeState, highSolver)
        If Not evaluated Or request.Problem.IsNumericalFailureState(probeState) Then
            If request.Problem.StoppedTerminally Then Exit Function
            If Not boundaryRefinementAttempted And Not lowSolver Is Nothing Then
                If lowSolver.Converged Then
                    ' На плато силовая проба может не сойтись, хотя сохранённая плоскость
                    ' уже позволяет найти первый предел. Инженерная задача отдельно
                    ' подтверждает равновесие всех усилий на исходной траектории нагрузки.
                    boundaryRefinementAttempted = True
                    request.Problem.TryFinalizeAfterSearchFailure request
                    If request.Problem.Converged Or request.Problem.StoppedTerminally Then Exit Function
                End If
            End If
            recoveryAttempts = recoveryAttempts + 1
            ' До подтвержденной физической скобки допуск lambda не завершает
            ' recovery: численный отказ не является физической верхней границей.
            ' Работу ограничивают бюджет и представимость следующей пробы.
            If request.Problem.MaxIterations > 0 Then
                If recoveryAttempts >= request.Problem.MaxIterations Then
                    request.Problem.SetNumericalFailure "Исчерпано допустимое число попыток восстановить поиск множителя нагрузки lambda. Равновесие следующей точки не подтверждено."
                    Exit Function
                End If
            End If
            reducedLambda = request.Problem.ReduceLambdaAfterNumericalFailure(lowLambda, highLambda, probeState)
            If Not IsFiniteNumber(reducedLambda) Or reducedLambda <= lowLambda Or reducedLambda >= highLambda Then
                request.Problem.SetNumericalFailure _
                    "Восстановление поиска lambda не уменьшило интервал; следующую пробу невозможно выбрать."
                Exit Function
            End If
            highLambda = reducedLambda
        ElseIf request.Problem.IsLimitState(probeState) Then
            FindBracket = True
            Exit Function
        ElseIf request.Problem.IsAcceptableState(probeState) Then
            lowLambda = highLambda ' Бюджет относится ко всему построению скобки: успешная проба не отменяет накопленные неудачи и не разрешает бесконечное recovery.
            Set lowSolver = highSolver
            request.Problem.RecordAcceptable loadPath, lowLambda
            If highLambda >= request.Problem.MaxLambda Then Exit Do
            If highLambda >= request.Problem.MaxLambda / 2# Then ' Доходим до самой MaxLambda без переполнения при удвоении и без пропуска конца интервала.
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
    ReportArithmeticOrContractFailure request.Problem, _
        "Расширение интервала предельного множителя нагрузки", Err.Number, Err.Description
End Function

' Переполнение при вычислении шага сохраняем как численную неудачу.
' Ошибка операции инженерной задачи считается внутренней; ранее полученную
' окончательную причину новым исключением не заменяем и расчёт не перезапускаем.
Private Sub ReportArithmeticOrContractFailure(ByVal context As ILimitSearchProblem, _
        ByVal stage As String, ByVal errorNumber As Long, ByVal description As String)
    If context.StoppedTerminally Then Exit Sub
    Dim failureCode As ESolverFailureCode
    Dim reason As String
    context.AppendDiagnostic "loadMultiplierRuntimeError=" & CStr(errorNumber) & "; stage=" & stage & _
        "; description=" & description
    If errorNumber = 6 Or errorNumber = 11 Then
        failureCode = sfcNumericalFailure
        reason = stage & ": не удалось вычислить численный шаг; число выходит за допустимый диапазон арифметики."
    Else
        failureCode = sfcInternalError
        reason = stage & ": внутренняя ошибка вычислений. Подробности сохранены в диагностическом журнале."
    End If
    ' Стандартную причину VBA уже пояснили по-русски. Подробную причину самой задачи
    ' сохраняем, поскольку она точнее общего сообщения. Сравнение текста
    ' меняет только вывод, а не код ошибки.
    Dim standardDescription As String
    Err.Clear ' Error$(n) при активном Err может вернуть явно заданный Description. Код и исходный текст уже сохранены аргументами и в журнале.
    If errorNumber >= 0 And errorNumber <= 65535 Then standardDescription = Error$(errorNumber)
    If Len(description) > 0 And StrComp(description, standardDescription, vbBinaryCompare) <> 0 Then
        reason = reason & " Причина: " & description
    End If
    context.SetFailure failureCode, reason
End Sub

' Запускает выбранный одномерный метод внутри уже найденной физической скобки.
Private Sub RunSearchInsideBracket(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lowLambda As Double, _
        ByVal highLambda As Double, ByVal lowSolver As CSectionSolver, _
        ByVal highSolver As CSectionSolver, ByVal methodName As String, _
        ByVal tolerance As Double, ByVal maxIterations As Long, ByRef pointAccepted As Boolean)
    Select Case methodName
        Case "Bisection"
            Dim iterations As Long
            request.Problem.AppendDiagnostic "searchMethod=Bisection"
            If SearchByBisectionCore(request, loadPath, lowLambda, lowSolver, highLambda, highSolver, _
                    tolerance, request.Problem.BisectionMaxIterations, iterations) Then
                pointAccepted = FinalizeBracketWithRecovery(request, loadPath, lowLambda, lowSolver, _
                    highLambda, highSolver, "Bisection", iterations, request.Problem.BisectionMaxIterations)
            End If
        Case "Brent"
            SearchByBrent request, loadPath, lowLambda, highLambda, tolerance, maxIterations, pointAccepted
        Case "Secant"
            SearchBySecant request, loadPath, lowLambda, highLambda, tolerance, maxIterations, pointAccepted
        Case Else
            request.Problem.SetInvalidInput _
                "Метод одномерного поиска должен быть Bisection, Brent или Secant."
    End Select
End Sub

' Малой ширины интервала недостаточно: нужно подтвердить физический предел.
' После численной неудачи сохраняем обе равновесные стороны и уточняем
' тот же интервал. Ошибки данных и программы не повторяем, лимит итераций не обнуляем.
Private Function FinalizeBracketWithRecovery(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lowLambda As Double, _
        ByVal lowSolver As CSectionSolver, ByVal highLambda As Double, _
        ByVal highSolver As CSectionSolver, ByVal methodName As String, _
        ByVal iterations As Long, ByVal maxIterations As Long) As Boolean
    Do
        If request.Problem.FinalizeBracket(request, loadPath, lowLambda, lowSolver, _
                highLambda, highSolver, methodName, iterations) Then
            FinalizeBracketWithRecovery = True
            Exit Function
        End If
        If request.Problem.Converged Or request.Problem.StoppedTerminally Then Exit Function
        If maxIterations > 0 And iterations >= maxIterations Then
            request.Problem.SetNumericalFailure _
                "Достигнуто допустимое число попыток поиска множителя нагрузки lambda. Точность интервала достигнута, но равновесие предельной точки не подтверждено."
            Exit Function
        End If
        Dim width As Double, recoveryTolerance As Double
        width = highLambda - lowLambda
        recoveryTolerance = width / 4#
        If recoveryTolerance <= 0# Or recoveryTolerance >= width Then
            request.Problem.SetNumericalFailure _
                "Уточнение предельной точки остановлено: интервал поиска множителя нагрузки lambda больше не удаётся уменьшить."
            Exit Function
        End If
        request.Problem.AppendDiagnostic "bracketFinalizationRecovery; width=" & _
            FormatNumberInvariant(width) & "; nextWidth=" & FormatNumberInvariant(recoveryTolerance) & _
            "; usedIterations=" & CStr(iterations)
        Dim remaining As Long, recoveryIterations As Long
        remaining = 0
        If maxIterations > 0 Then remaining = maxIterations - iterations
        If Not SearchByBisectionCore(request, loadPath, lowLambda, lowSolver, highLambda, _
                highSolver, recoveryTolerance, remaining, recoveryIterations) Then Exit Function
        iterations = iterations + recoveryIterations
    Loop
End Function

' Уточняет найденный интервал lambda делением пополам. Несущая способность
' и образование трещины используют общий порядок проб;
' их физический критерий и принимаемая сторона интервала задаются самой проверкой.
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
        Dim evaluated As Boolean
        evaluated = EvaluateBisectionPoint(request, loadPath, startSolver, midLambda, _
            midResidual, probeState, midSolver)
        If Not request.Problem.StoppedTerminally Then
            If Not evaluated Or IsNumericalFailureState(request, probeState) Then
                If Not highSolver Is Nothing Then
                    request.Problem.AppendDiagnostic "bisectionRecovery=oppositeEquilibriumStart; lambda=" & FormatNumberInvariant(midLambda)
                    evaluated = EvaluateBisectionPoint(request, loadPath, highSolver, midLambda, _
                        midResidual, probeState, midSolver)
                End If
            End If
        End If
        If Not evaluated Then
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
        ByVal highLambda As Double, ByVal tolerance As Double, ByVal maxIterations As Long, _
        ByRef pointAccepted As Boolean)
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
        If FinalizeKnownProbe(request, loadPath, b, stateB, "Brent", solverB, 0) Then
            pointAccepted = True
            Exit Sub
        End If
        If request.Problem.StoppedTerminally Then Exit Sub
    End If
    If fa >= 0# Or fb <= 0# Then
        request.Problem.SetInvalidInput _
            "Для метода Brent нужен интервал поиска: при нижнем множителе нагрузки lambda состояние должно быть допустимым, при верхнем — достигать или превышать физический предел."
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
            If FinalizeKnownProbe(request, loadPath, b, stateB, "Brent", solverB, iter) Then
                pointAccepted = True
                Exit Sub
            End If
            If request.Problem.StoppedTerminally Then Exit Sub
        End If
        If Abs(c - b) <= tolerance Then
            ' Метод Brent выбирает точку с ближайшим к нулю критерием, которая может быть
            ' с любой стороны предела. Передаём обе стороны для окончательной проверки:
            ' несущая способность принимает нижнюю, образование трещины подтверждает верхнюю.
            If IsAcceptableState(request, stateB) Then
                pointAccepted = FinalizeBracketWithRecovery(request, loadPath, b, solverB, c, solverC, "Brent", iter, maxIterations)
            ElseIf IsAcceptableState(request, stateC) Then
                pointAccepted = FinalizeBracketWithRecovery(request, loadPath, c, solverC, b, solverB, "Brent", iter, maxIterations)
            Else
                request.Problem.SetNumericalFailure "В итоговом интервале поиска Brent нет сохранённой точки с подтверждённым допустимым равновесием."
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
        ByVal highLambda As Double, ByVal tolerance As Double, ByVal maxIterations As Long, _
        ByRef pointAccepted As Boolean)
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
        If FinalizeKnownProbe(request, loadPath, x1, state1, "Secant", solver1, 0) Then
            pointAccepted = True
            Exit Sub
        End If
        If request.Problem.StoppedTerminally Then Exit Sub
    End If
    If f0 >= 0# Or f1 <= 0# Then
        request.Problem.SetInvalidInput _
            "Для метода Secant нужен интервал поиска: при нижнем множителе нагрузки lambda состояние должно быть допустимым, при верхнем — достигать или превышать физический предел."
        Exit Sub
    End If
    bracketLow = x0
    bracketHigh = x1
    Set bracketLowSolver = solver0
    Set bracketHighSolver = solver1

    For iter = 1 To maxIterations
        If Abs(f1) <= tolerance And CanFinalizeSecantProbe(request, state1, solver1) Then
            If FinalizeKnownProbe(request, loadPath, x1, state1, "Secant", solver1, iter) Then
                pointAccepted = True
                Exit Sub
            End If
            If request.Problem.StoppedTerminally Then Exit Sub
        End If
        If Abs(bracketHigh - bracketLow) <= tolerance Then
            pointAccepted = FinalizeBracketWithRecovery(request, loadPath, bracketLow, bracketLowSolver, _
                bracketHigh, bracketHighSolver, "Secant", iter, maxIterations)
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
            If FinalizeKnownProbe(request, loadPath, x2, state2, "Secant", solver2, iter) Then
                pointAccepted = True
                Exit Sub
            End If
            If request.Problem.StoppedTerminally Then Exit Sub
        End If
        If Abs(bracketHigh - bracketLow) <= tolerance Then
            pointAccepted = FinalizeBracketWithRecovery(request, loadPath, bracketLow, bracketLowSolver, _
                bracketHigh, bracketHighSolver, "Secant", iter, maxIterations)
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

' Если после исчерпания делений пополам интервал всё ещё шире допуска,
' сохраняет численную неудачу. Недостаточно уточнённая нижняя нагрузка
' не должна выдаваться за найденный предел.
Private Sub HandleBisectionIterationLimit(ByVal request As CLimitSearchRequest, _
        ByVal maxIterations As Long, ByVal remainingWidth As Double, ByVal tolerance As Double)
    request.Problem.HandleIterationLimit maxIterations, remainingWidth, tolerance
End Sub

' Повторно оценивает то же lambda, если задача не передала соответствующее НДС.
' Возвращает фактический итог физической проверки, не меняя множитель нагрузки.
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

' Передаёт уже рассчитанную пробную нагрузку на подтверждение без нового НДС.
' Несошедшаяся проба физического предела не доказывает. При отказе без
' окончательной причины остановки общий поиск продолжает уточнение интервала.
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

' Выбирает НДС для начального приближения следующей пробы деления пополам.
' Сначала использует равновесие допустимой стороны, затем подтверждённой верхней.
' Пригодность старта проверяет сама инженерная задача.
Private Function PreferredBisectionStartSolver(ByVal lowSolver As CSectionSolver, _
        ByVal highSolver As CSectionSolver) As CSectionSolver
    If lowSolver Is Nothing Then
        Set PreferredBisectionStartSolver = highSolver
    Else
        Set PreferredBisectionStartSolver = lowSolver
    End If
End Function

' Оценивает одну нагрузку lambda для деления интервала пополам.
' Получает критерий и соответствующее НДС для следующих стартов
' и окончательного подтверждения без повторного решения.
Private Function EvaluateBisectionPoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByVal lambdaValue As Double, ByRef functionValue As Double, _
        ByRef probeState As String, ByRef solver As CSectionSolver) As Boolean
    EvaluateBisectionPoint = request.Problem.EvaluatePoint(request, loadPath, _
        startSolver, lambdaValue, functionValue, probeState, solver)
End Function

' Передаёт инженерной задаче причину неудачной нагрузки lambda.
' Уже записанная окончательная причина сохраняется; численная неудача
' означает отсутствие подтверждённого равновесия этой пробы.
Private Sub HandleBisectionEvaluationFailure(ByVal request As CLimitSearchRequest, _
        ByVal lambdaValue As Double)
    If request.Problem.StoppedTerminally Then Exit Sub
    request.Problem.SetNumericalFailure _
        "Не удалось найти равновесие при lambda = " & _
        FormatNumberInvariant(lambdaValue) & " в поиске предельной точки."
End Sub

' Обрабатывает численную неудачу, которую задача вернула отдельным статусом NumericalFailure.
Private Sub HandleBisectionNumericalFailure(ByVal request As CLimitSearchRequest, _
        ByVal lowLambda As Double, ByVal failedLambda As Double, ByVal probeState As String)
    request.Problem.HandleBracketNumericalFailure lowLambda, failedLambda, probeState
End Sub

' Останавливает деление пополам при неизвестном статусе пробы.
' Это неверный ответ программного объекта задачи, а не физический результат.
Private Sub HandleBisectionUnexpectedState(ByVal request As CLimitSearchRequest, _
        ByVal probeState As String)
    request.Problem.HandleUnexpectedState probeState
End Sub

' Оценивает нагрузку и возвращает её НДС. Brent и Secant сохраняют
' это НДС вместе с соответствующей стороной интервала,
' чтобы не принять состояние от другой пробной нагрузки.
Private Function EvaluateLimitPoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByRef functionValue As Double, ByRef probeState As String, _
        ByRef solver As CSectionSolver) As Boolean
    EvaluateLimitPoint = request.Problem.EvaluatePoint(request, loadPath, Nothing, _
        lambdaValue, functionValue, probeState, solver)
    If EvaluateLimitPoint Then
        If IsNumericalFailureState(request, probeState) Then
            HandleBisectionEvaluationFailure request, lambdaValue ' Неудачная проба может вернуть нулевую невязку-заглушку. Brent/Secant не должны принимать ее за корень или менять знак скобки.
            EvaluateLimitPoint = False
        End If
    End If
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

' Проверяет пригодность пробы Secant для итогового НДС. Допустимая точка
' уже находится с безопасной стороны; для предельной точки требуется
' подтверждённое равновесие. Несошедшаяся проба требует отдельного уточнения.
Private Function CanFinalizeSecantProbe(ByVal request As CLimitSearchRequest, _
        ByVal probeState As String, ByVal solver As CSectionSolver) As Boolean
    If IsAcceptableState(request, probeState) Then
        CanFinalizeSecantProbe = True
    ElseIf IsLimitState(request, probeState) Then
        If solver Is Nothing Then
            CanFinalizeSecantProbe = True ' Аналитическая контрольная задача может подтвердить критерий без объекта НДС. Для реального сечения равновесие и физические пределы проверяет её FinalizeAt.
        Else
            CanFinalizeSecantProbe = solver.Converged
        End If
    End If
End Function

' Ограничивает следующее lambda метода Secant найденным интервалом
' между допустимой и предельной нагрузками.
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
