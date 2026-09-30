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

Private Const STATE_ACCEPTABLE As String = "Acceptable" ' Lambda-точка еще не достигла физического предела.
Private Const STATE_CRACK_LIMIT As String = "CrackFormationLimit" ' Для CrackFormation eps_bt,ult уже достигнута.
Private Const STATE_NUMERICAL_FAILURE As String = "NumericalFailure" ' Численная неудача пробной lambda-точки.
Private Const STATE_INVALID_INPUT As String = "InvalidInput" ' Ошибка исходных данных или настройки search.

' Выполняет LoadMultiplier для несущей способности через общий алгоритм lambda-search.
Public Function ExecuteCapacity(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request Is Nothing Then Err.Raise vbObjectError + 4220, "CLoadMultiplierSearch", "Не передан запрос limit-search."

    If Not request.Problem.Prepare(request) Then
        Set ExecuteCapacity = request.Problem.BuildResult(request, False)
        Exit Function
    End If

    Dim loadPath As CLoadPathVector
    Set loadPath = request.Problem.BuildLoadPath(request)

    RunCapacitySearch request, loadPath
    Set ExecuteCapacity = request.Problem.BuildResult(request, False)
End Function

' Выполняет LoadMultiplier для точки образования нормальной трещины. Crack-слой
' задает критерий eps_bt,ult и load-path, а bracket/bisection остаются общими.
Public Function ExecuteCrackFormation(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request Is Nothing Then Err.Raise vbObjectError + 4221, "CLoadMultiplierSearch", "Не передан запрос limit-search."

    request.Problem.AppendDiagnostic "LoadMultiplier: начинаем подбор lambda_crc пробными расчетами НДС; путь " & _
        request.Problem.FormationPathDisplayName() & "."

    Dim loadPath As CLoadPathVector
    Set loadPath = request.Problem.BuildLoadPath(request)

    Dim searchSucceeded As Boolean
    searchSucceeded = RunCrackFormationSearch(request, loadPath)
    Set ExecuteCrackFormation = request.Problem.BuildResult(request, searchSucceeded)
End Function

' Запускает общий поиск для capacity: нижняя lambda=0 должна быть acceptable,
' верхняя граница расширяется до первого физического предела.
Private Sub RunCapacitySearch(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector)
    Dim lowLambda As Double
    Dim highLambda As Double
    Dim probeState As String
    Dim functionValue As Double
    lowLambda = 0#
    highLambda = request.Problem.InitialLambdaStep
    If highLambda <= 0# Then highLambda = 1#

    If Not EvaluateLimitFunction(request, loadPath, lowLambda, functionValue, probeState) Then Exit Sub
    If Not IsAcceptableState(request, probeState) Then
        request.Problem.HandleInitialFailure probeState
        Exit Sub
    End If
    request.Problem.RecordAcceptable loadPath, lowLambda

    If Not FindCapacityBracket(request, loadPath, lowLambda, highLambda) Then
        If request.ForceOnly Then request.Problem.AcceptRetainedCapacity request
        Exit Sub
    End If

    RunSearchInsideBracket request, loadPath, lowLambda, highLambda, _
        request.Problem.SearchMethod, request.Problem.LambdaTolerance, request.Problem.MaxIterations
    If request.ForceOnly Then request.Problem.AcceptRetainedCapacity request
End Sub

' Ищет верхнюю физическую границу capacity так же, как прежний
' CCapacitySolver: acceptable-точки сдвигают low, NumericalFailure уменьшает
' верхнюю lambda, а физический предел завершает скобку.
Private Function FindCapacityBracket(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByRef lowLambda As Double, _
        ByRef highLambda As Double) As Boolean
    Dim probeState As String
    Dim functionValue As Double
    Do While highLambda <= request.Problem.MaxLambda
        If Not EvaluateLimitFunction(request, loadPath, highLambda, functionValue, probeState) Then Exit Function

        If IsAcceptableState(request, probeState) Then
            lowLambda = highLambda
            request.Problem.RecordAcceptable loadPath, highLambda
            highLambda = highLambda * 2#
        ElseIf IsLimitState(request, probeState) Then
            FindCapacityBracket = True
            Exit Function
        ElseIf IsNumericalFailureState(request, probeState) Then
            highLambda = request.Problem.ReduceLambdaAfterNumericalFailure(lowLambda, highLambda, probeState)
            If highLambda <= lowLambda Then Exit Function
        Else
            request.Problem.SetNumericalFailure "Поиск несущей способности остановлен: " & probeState & "."
            Exit Function
        End If
    Loop

    request.Problem.HandleBracketNotFound lowLambda
End Function

' Запускает LoadMultiplier для CrackFormation. Начальная точка lambda=0 может
' быть не обычным solve, а нулевой нагрузкой; это обрабатывает crack-критерий.
Private Function RunCrackFormationSearch(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector) As Boolean
    Dim lowLambda As Double
    Dim highLambda As Double
    Dim lowSolver As CSectionSolver
    Dim highSolver As CSectionSolver
    Dim lowResidual As Double
    Dim highResidual As Double
    Dim probeState As String
    lowLambda = 0#
    highLambda = 1#

    If Not EvaluateCrackFunction(request, loadPath, Nothing, lowLambda, lowResidual, probeState, lowSolver) Then
        request.Problem.HandleInitialFailure probeState
        Exit Function
    End If

    If StrComp(request.CrackFormationPath, "LambdaN", vbTextCompare) = 0 And request.TargetN <= 0.000001 Then
        If lowResidual >= 0# Then
            request.Problem.MarkFormationPathFallback "Трещина уже образуется от постоянного момента выбранного пути " & _
                request.Problem.FormationPathDisplayName() & "; сжимающая или нулевая N не масштабируется для поиска Ncrc, psi_s остается равным 1."
        Else
            request.Problem.FinishNoCrack "Трещина не образована от постоянного момента, а выбранный путь " & _
                request.Problem.FormationPathDisplayName() & " не масштабирует растягивающую продольную силу."
        End If
        RunCrackFormationSearch = True
        Exit Function
    End If

    If lowResidual >= 0# Then
        request.Problem.MarkFormationPathFallback "Трещина уже образуется в постоянной части выбранного пути " & _
            request.Problem.FormationPathDisplayName() & "; точка sigma_s,crc по lambda не определяется, psi_s остается равным 1."
        RunCrackFormationSearch = True
        Exit Function
    End If

    If Not request.Problem.HasScaledFormationLoad(request) Then
        request.Problem.FinishNoCrack "Трещина не образована в постоянной части, а выбранный путь " & _
            request.Problem.FormationPathDisplayName() & " не содержит ненулевой масштабируемой нагрузки."
        RunCrackFormationSearch = True
        Exit Function
    End If

    If Not FindCrackFormationBracket(request, loadPath, lowLambda, lowSolver, lowResidual, _
            highLambda, highSolver, highResidual) Then Exit Function

    If highResidual < 0# Then
        request.Problem.HandleBracketNotFound highLambda
        Exit Function
    End If

    Dim acceptedLambda As Double
    Dim acceptedSolver As CSectionSolver
    Dim iterations As Long
    If SearchCrackByBisection(request, loadPath, lowLambda, lowSolver, highLambda, highSolver, _
            acceptedLambda, acceptedSolver, iterations) Then
        request.Problem.FinalizeAt request, loadPath, acceptedLambda, vbNullString, _
            "Bisection", acceptedSolver, iterations
        RunCrackFormationSearch = True
    End If
End Function

' Ищет верхнюю точку crack-скобки. Если перегруженная точка не решается,
' верхняя граница сдвигается вниз, чтобы подойти к eps_bt,ult с рабочей стороны.
Private Function FindCrackFormationBracket(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByRef lowerLambda As Double, _
        ByRef lowerSolver As CSectionSolver, ByRef lowerResidual As Double, _
        ByRef upperLambda As Double, ByRef upperSolver As CSectionSolver, _
        ByRef upperResidual As Double) As Boolean
    Dim crack As CCrackWidthCalculator
    Set crack = request.CrackCalculator

    Dim probeState As String
    If EvaluateCrackFunction(request, loadPath, lowerSolver, upperLambda, upperResidual, probeState, upperSolver) Then
        GoTo UpperSolved
    End If

    If TryRecoverCrackBracketBeforeFailedUpper(request, loadPath, lowerLambda, lowerSolver, lowerResidual, _
            upperLambda, upperSolver, upperResidual) Then
        FindCrackFormationBracket = True
        Exit Function
    End If

    crack.LimitSearchSetCrackFormationStopReason _
        "Не удалось подобрать сходящуюся точку образования трещины до перегруженной точки lambda = " & _
        FormatNumberInvariant(upperLambda) & "."
    Exit Function

UpperSolved:
    Do While upperResidual < 0# And upperLambda < crack.LimitSearchCrackBracketLimit
        lowerLambda = upperLambda
        Set lowerSolver = upperSolver
        lowerResidual = upperResidual
        upperLambda = MinDouble(upperLambda * 2#, crack.LimitSearchCrackBracketLimit)
        If Not EvaluateCrackFunction(request, loadPath, lowerSolver, upperLambda, upperResidual, probeState, upperSolver) Then
            If TryRecoverCrackBracketBeforeFailedUpper(request, loadPath, lowerLambda, lowerSolver, lowerResidual, _
                    upperLambda, upperSolver, upperResidual) Then
                FindCrackFormationBracket = True
                Exit Function
            End If
            crack.LimitSearchSetCrackFormationStopReason _
                "Не удалось найти равновесие при lambda = " & FormatNumberInvariant(upperLambda) & _
                " в поиске трещинообразования."
            Exit Function
        End If
    Loop
    FindCrackFormationBracket = True
End Function

' Если верхняя crack-проба не сошлась, ищет рабочую точку с образованием
' трещины между последней сходящейся нижней точкой и перегруженной lambda.
' Это не меняет критерий трещинообразования, а только не дает перегруженной
' пробе преждевременно превращать весь поиск в NumFail.
Private Function TryRecoverCrackBracketBeforeFailedUpper(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByRef lowerLambda As Double, _
        ByRef lowerSolver As CSectionSolver, ByRef lowerResidual As Double, _
        ByRef upperLambda As Double, ByRef upperSolver As CSectionSolver, _
        ByRef upperResidual As Double) As Boolean
    Dim crack As CCrackWidthCalculator
    Set crack = request.CrackCalculator

    Dim failedLambda As Double
    failedLambda = upperLambda

    Dim iter As Long
    For iter = 1 To crack.LimitSearchCrackMaxIterations
        If failedLambda - lowerLambda <= crack.LimitSearchCrackLambdaTolerance * MaxDouble(1#, failedLambda) Then Exit For

        Dim midLambda As Double
        Dim midSolver As CSectionSolver
        Dim midResidual As Double
        Dim probeState As String
        midLambda = 0.5 * (lowerLambda + failedLambda)
        If EvaluateCrackFunction(request, loadPath, lowerSolver, midLambda, midResidual, probeState, midSolver) Then
            If midResidual >= 0# Then
                upperLambda = midLambda
                Set upperSolver = midSolver
                upperResidual = midResidual
                TryRecoverCrackBracketBeforeFailedUpper = True
                Exit Function
            End If
            lowerLambda = midLambda
            Set lowerSolver = midSolver
            lowerResidual = midResidual
        Else
            failedLambda = midLambda
        End If
    Next iter

    If Not lowerSolver Is Nothing Then
        crack.LimitSearchAppendDiagnostic _
            "Поиск трещинообразования: последняя сходящаяся точка перед перегрузкой lambda=" & _
            FormatNumberInvariant(lowerLambda) & "; невязка=" & FormatNumberInvariant(lowerResidual)
    End If
End Function

' Для CrackFormation используется тот же bisection-цикл, что и для Capacity.
' Отличается только критерий одной lambda-точки и финализация найденного состояния.
Private Function SearchCrackByBisection(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lowerLambda As Double, _
        ByVal lowerSolver As CSectionSolver, ByVal upperLambda As Double, _
        ByVal upperSolver As CSectionSolver, ByRef acceptedLambda As Double, _
        ByRef acceptedSolver As CSectionSolver, ByRef iterations As Long) As Boolean
    Dim crack As CCrackWidthCalculator
    Set crack = request.CrackCalculator

    If Not SearchByBisectionCore(request, loadPath, lowerLambda, lowerSolver, _
            upperLambda, upperSolver, crack.LimitSearchCrackLambdaTolerance, _
            crack.LimitSearchCrackMaxIterations, iterations) Then Exit Function

    acceptedLambda = upperLambda
    Set acceptedSolver = upperSolver
    SearchCrackByBisection = True
End Function

' Запускает выбранный одномерный метод внутри уже найденной физической скобки.
Private Sub RunSearchInsideBracket(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lowLambda As Double, _
        ByVal highLambda As Double, ByVal methodName As String, _
        ByVal tolerance As Double, ByVal maxIterations As Long)
    Select Case methodName
        Case "Bisection"
            SearchByBisection request, loadPath, lowLambda, highLambda, tolerance
        Case "Brent"
            SearchByBrent request, loadPath, lowLambda, highLambda, tolerance, maxIterations
        Case "Secant"
            SearchBySecant request, loadPath, lowLambda, highLambda, tolerance, maxIterations
        Case Else
            request.CapacitySolver.LimitSearchSetCapacityInvalidInput _
                "Capacity.SearchMethod должен быть Bisection, Brent или Secant."
    End Select
End Sub

' Устойчивый базовый метод: делит физическую скобку пополам.
Private Sub SearchByBisection(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lowLambda As Double, _
        ByVal highLambda As Double, ByVal tolerance As Double)
    request.CapacitySolver.LimitSearchAppendCapacityDiagnostic "searchMethod=Bisection"

    Dim lowSolver As CSectionSolver
    Dim highSolver As CSectionSolver
    Dim iterations As Long
    If Not SearchByBisectionCore(request, loadPath, lowLambda, lowSolver, _
            highLambda, highSolver, tolerance, 0, iterations) Then Exit Sub

    request.CapacitySolver.LimitSearchFinalizeCapacityFromRetainedAcceptable _
        lowLambda, loadPath, "Bisection"
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

        midLambda = (lowLambda + highLambda) / 2#
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
    request.CapacitySolver.LimitSearchAppendCapacityDiagnostic "searchMethod=Brent"

    Dim a As Double, b As Double, c As Double
    Dim fa As Double, fb As Double, fc As Double
    Dim stateA As String, stateB As String, stateC As String
    Dim d As Double, e As Double
    Dim iter As Long
    a = lowLambda
    b = highLambda
    If Not EvaluateLimitFunction(request, loadPath, a, fa, stateA) Then Exit Sub
    If Not EvaluateLimitFunction(request, loadPath, b, fb, stateB) Then Exit Sub
    If fa >= 0# Or fb <= 0# Then
        request.CapacitySolver.LimitSearchSetCapacityInvalidInput _
            "Методу Brent нужна скобка: нижняя lambda с запасом и верхняя lambda за физическим пределом."
        Exit Sub
    End If

    c = a
    fc = fa
    stateC = stateA
    d = b - a
    e = d

    For iter = 1 To maxIterations
        If fb * fc > 0# Then
            c = a
            fc = fa
            stateC = stateA
            d = b - a
            e = d
        End If
        If Abs(fc) < Abs(fb) Then
            a = b: b = c: c = a
            fa = fb: fb = fc: fc = fa
            stateA = stateB: stateB = stateC: stateC = stateA
        End If

        Dim tol As Double
        Dim xm As Double
        tol = MaxDouble(tolerance / 2#, 0.000000000001 * Abs(b))
        xm = (c - b) / 2#
        If Abs(xm) <= tol Then
            FinalizeAfterProbe request, loadPath, b, "Brent"
            Exit Sub
        End If

        If Abs(e) >= tol And Abs(fa) > Abs(fb) Then
            Dim s As Double, p As Double, q As Double
            Dim r As Double
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
            If p > 0# Then q = -q
            p = Abs(p)
            If 2# * p < MinDouble(3# * xm * q - Abs(tol * q), Abs(e * q)) Then
                e = d
                d = p / q
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
        If Abs(d) > tol Then
            b = b + d
        Else
            b = b + Sgn(xm) * tol
        End If
        If b <= lowLambda Or b >= highLambda Or Not IsFiniteNumber(b) Then
            request.CapacitySolver.LimitSearchSetCapacityNumericalFailure _
                "Пробная точка Brent вышла за физическую скобку."
            Exit Sub
        End If
        If Not EvaluateLimitFunction(request, loadPath, b, fb, stateB) Then Exit Sub
    Next iter

    request.CapacitySolver.LimitSearchSetCapacityNumericalFailure "Поиск Brent достиг лимита итераций."
End Sub

' Метод секущих: строит следующую lambda по двум предыдущим значениям функции,
' но дополнительно удерживает точку внутри физической скобки.
Private Sub SearchBySecant(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lowLambda As Double, _
        ByVal highLambda As Double, ByVal tolerance As Double, ByVal maxIterations As Long)
    request.CapacitySolver.LimitSearchAppendCapacityDiagnostic "searchMethod=Secant"

    Dim x0 As Double, x1 As Double, x2 As Double
    Dim f0 As Double, f1 As Double, f2 As Double
    Dim state0 As String, state1 As String, state2 As String
    Dim bracketLow As Double
    Dim bracketHigh As Double
    Dim iter As Long

    x0 = lowLambda
    x1 = highLambda
    If Not EvaluateLimitFunction(request, loadPath, x0, f0, state0) Then Exit Sub
    If Not EvaluateLimitFunction(request, loadPath, x1, f1, state1) Then Exit Sub
    If f0 >= 0# Or f1 <= 0# Then
        request.CapacitySolver.LimitSearchSetCapacityInvalidInput _
            "Методу Secant нужна скобка: нижняя lambda с запасом и верхняя lambda за физическим пределом."
        Exit Sub
    End If
    bracketLow = x0
    bracketHigh = x1

    For iter = 1 To maxIterations
        If Abs(f1) <= tolerance Then
            FinalizeAfterProbe request, loadPath, x1, "Secant"
            Exit Sub
        End If
        If Abs(bracketHigh - bracketLow) <= tolerance Then
            FinalizeAfterProbe request, loadPath, bracketLow, "Secant"
            Exit Sub
        End If
        If Abs(f1 - f0) <= 0.000000000001 Then
            request.CapacitySolver.LimitSearchSetCapacityNumericalFailure _
                "Поиск Secant остановлен: значения целевой функции слишком близки."
            Exit Sub
        End If

        x2 = x1 - f1 * (x1 - x0) / (f1 - f0)
        If Not IsFiniteNumber(x2) Then
            request.CapacitySolver.LimitSearchSetCapacityNumericalFailure _
                "Поиск Secant получил NaN или переполнение."
            Exit Sub
        End If
        x2 = GuardSecantCandidate(request, x2, x1, bracketLow, bracketHigh)
        If x2 <= bracketLow Or x2 >= bracketHigh Or Not IsFiniteNumber(x2) Then
            request.CapacitySolver.LimitSearchSetCapacityNumericalFailure _
                "Пробная точка Secant вышла за физическую скобку."
            Exit Sub
        End If
        If Abs(x2 - x1) <= 0.000000000001 Then
            request.CapacitySolver.LimitSearchSetCapacityNumericalFailure _
                "Поиск Secant застрял: новая lambda почти не отличается от предыдущей."
            Exit Sub
        End If

        If Not EvaluateLimitFunction(request, loadPath, x2, f2, state2) Then Exit Sub
        UpdateSecantBracket request, x2, f2, state2, bracketLow, bracketHigh
        If Abs(f2) <= tolerance Then
            FinalizeAfterProbe request, loadPath, x2, "Secant"
            Exit Sub
        End If
        If Abs(bracketHigh - bracketLow) <= tolerance Then
            FinalizeAfterProbe request, loadPath, bracketLow, "Secant"
            Exit Sub
        End If

        x0 = x1
        f0 = f1
        state0 = state1
        x1 = x2
        f1 = f2
        state1 = state2
    Next iter

    request.CapacitySolver.LimitSearchSetCapacityNumericalFailure "Поиск Secant достиг лимита итераций."
End Sub

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

    If request.ResultKind = rkCapacity Then
        request.CapacitySolver.LimitSearchSetCapacityNumericalFailure reason
    Else
        request.CrackCalculator.LimitSearchSetCrackFormationStopReason reason
    End If
End Sub

' Пересчитывает итоговую lambda и передает доменному объекту финализацию.
Private Sub FinalizeAfterProbe(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal methodName As String)
    Dim functionValue As Double
    Dim probeState As String
    If Not EvaluateLimitFunction(request, loadPath, lambdaValue, functionValue, probeState) Then Exit Sub
    request.CapacitySolver.LimitSearchFinalizeCapacityLoadMultiplier lambdaValue, loadPath, probeState, methodName
End Sub

' Выбирает стартовое НДС для следующей bisection-пробы.
' Для CrackFormation выгодно стартовать от последней сходящейся нижней точки; если ее нет,
' используется верхняя точка. Для Capacity solver здесь не нужен, поэтому возвращается Nothing.
Private Function PreferredBisectionStartSolver(ByVal lowSolver As CSectionSolver, _
        ByVal highSolver As CSectionSolver) As CSectionSolver
    If lowSolver Is Nothing Then
        Set PreferredBisectionStartSolver = highSolver
    Else
        Set PreferredBisectionStartSolver = lowSolver
    End If
End Function

' Оценивает одну lambda-точку для общего bisection-цикла.
' Capacity возвращает только значение функции и статус, а CrackFormation дополнительно
' возвращает найденный solver, чтобы финализировать PreCrackState без повторного solve.
Private Function EvaluateBisectionPoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByVal lambdaValue As Double, ByRef functionValue As Double, _
        ByRef probeState As String, ByRef solver As CSectionSolver) As Boolean
    If request.ResultKind = rkCapacity Then
        Set solver = Nothing
        EvaluateBisectionPoint = EvaluateLimitFunction(request, loadPath, _
            lambdaValue, functionValue, probeState)
    Else
        EvaluateBisectionPoint = EvaluateCrackFunction(request, loadPath, _
            startSolver, lambdaValue, functionValue, probeState, solver)
    End If
End Function

' Возвращает доменному объекту понятную причину отказа пробной lambda-точки.
' Для Capacity callback обычно уже записывает статус; для CrackFormation фиксируем stop-reason явно.
Private Sub HandleBisectionEvaluationFailure(ByVal request As CLimitSearchRequest, _
        ByVal lambdaValue As Double)
    If request.ResultKind = rkCapacity Then Exit Sub
    request.CrackCalculator.LimitSearchSetCrackFormationStopReason _
        "Не удалось найти равновесие при lambda = " & _
        FormatNumberInvariant(lambdaValue) & " в поиске трещинообразования."
End Sub

' Обрабатывает статус NumericalFailure, если callback смог классифицировать точку без обычного False.
Private Sub HandleBisectionNumericalFailure(ByVal request As CLimitSearchRequest, _
        ByVal lowLambda As Double, ByVal failedLambda As Double, ByVal probeState As String)
    If request.ResultKind = rkCapacity Then
        request.CapacitySolver.LimitSearchHandleCapacityBracketNumericalFailure lowLambda, probeState
    Else
        request.CrackCalculator.LimitSearchSetCrackFormationStopReason _
            "Не удалось найти равновесие при lambda = " & _
            FormatNumberInvariant(failedLambda) & " в поиске трещинообразования."
    End If
End Sub

' Защищает общий bisection от неизвестного статуса пробной точки.
' Такая ситуация означает не физический результат, а нарушение контракта callback-а.
Private Sub HandleBisectionUnexpectedState(ByVal request As CLimitSearchRequest, _
        ByVal probeState As String)
    If request.ResultKind = rkCapacity Then
        request.CapacitySolver.LimitSearchSetCapacityNumericalFailure _
            "Поиск Bisection остановлен: " & probeState & "."
    Else
        request.CrackCalculator.LimitSearchSetCrackFormationStopReason _
            "Поиск трещинообразования остановлен: " & probeState & "."
    End If
End Sub

' Единая целевая функция для capacity. CrackFormation использует отдельный
' wrapper, потому что ему нужно вернуть solver найденной верхней точки.
Private Function EvaluateLimitFunction(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByRef functionValue As Double, ByRef probeState As String) As Boolean
    EvaluateLimitFunction = request.CapacitySolver.LimitSearchEvaluateCapacityLoadMultiplier( _
        request.Section, request.ConcreteMaterial, request.SteelMaterial, _
        loadPath, lambdaValue, functionValue, probeState)
End Function

' Оценивает crack lambda-точку и возвращает solver, пригодный для последующего
' warm-start или сохранения PreCrackState.
Private Function EvaluateCrackFunction(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByVal lambdaValue As Double, ByRef functionValue As Double, _
        ByRef probeState As String, ByRef solver As CSectionSolver) As Boolean
    Dim concreteMaterial As CMaterialDiagram
    Dim steelMaterial As CMaterialDiagram
    Set concreteMaterial = request.ConcreteMaterial
    Set steelMaterial = request.SteelMaterial

    EvaluateCrackFunction = request.CrackCalculator.LimitSearchEvaluateCrackFormationLoadMultiplier( _
        request.Section, concreteMaterial, steelMaterial, _
        loadPath, startSolver, lambdaValue, functionValue, probeState, solver)
End Function

' Проверяет, что lambda-точка находится ниже предела.
Private Function IsAcceptableState(ByVal request As CLimitSearchRequest, ByVal probeState As String) As Boolean
    If request.ResultKind = rkCapacity Then
        IsAcceptableState = request.CapacitySolver.LimitSearchCapacityStateIsAcceptable(probeState)
    Else
        IsAcceptableState = (probeState = STATE_ACCEPTABLE)
    End If
End Function

' Проверяет, что lambda-точка находится на физическом пределе.
Private Function IsLimitState(ByVal request As CLimitSearchRequest, ByVal probeState As String) As Boolean
    If request.ResultKind = rkCapacity Then
        IsLimitState = request.CapacitySolver.LimitSearchCapacityStateIsPhysicalLimit(probeState)
    Else
        IsLimitState = (probeState = STATE_CRACK_LIMIT)
    End If
End Function

' Проверяет, что отказ lambda-точки является численной неудачей.
Private Function IsNumericalFailureState(ByVal request As CLimitSearchRequest, ByVal probeState As String) As Boolean
    If request.ResultKind = rkCapacity Then
        IsNumericalFailureState = request.CapacitySolver.LimitSearchCapacityStateIsNumericalFailure(probeState)
    Else
        IsNumericalFailureState = (probeState = STATE_NUMERICAL_FAILURE)
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
            GuardSecantCandidate = (currentLambda + bracketLow) / 2#
        Else
            GuardSecantCandidate = bracketLow + 0.25 * (bracketHigh - bracketLow)
        End If
    ElseIf rawCandidate >= bracketHigh Then
        If currentLambda > bracketLow And currentLambda < bracketHigh Then
            GuardSecantCandidate = (currentLambda + bracketHigh) / 2#
        Else
            GuardSecantCandidate = bracketLow + 0.75 * (bracketHigh - bracketLow)
        End If
    End If

    request.CapacitySolver.LimitSearchAppendCapacityDiagnostic _
        "secantGuard=insideBracket; rawLambda=" & FormatNumberInvariant(rawCandidate) & _
        "; guardedLambda=" & FormatNumberInvariant(GuardSecantCandidate) & _
        "; bracketLow=" & FormatNumberInvariant(bracketLow) & _
        "; bracketHigh=" & FormatNumberInvariant(bracketHigh)
End Function

' Обновляет защитную физическую скобку Secant.
Private Sub UpdateSecantBracket(ByVal request As CLimitSearchRequest, _
        ByVal lambdaValue As Double, ByVal functionValue As Double, _
        ByVal probeState As String, ByRef bracketLow As Double, ByRef bracketHigh As Double)
    If functionValue < 0# Then
        If lambdaValue > bracketLow Then bracketLow = lambdaValue
    Else
        If lambdaValue < bracketHigh Then bracketHigh = lambdaValue
    End If
End Sub

' Проверяет, что число конечно и пригодно для дальнейшей арифметики.
Private Function IsFiniteNumber(ByVal value As Double) As Boolean
    If value <> value Then Exit Function
    If Abs(value) > 1E+100 Then Exit Function
    IsFiniteNumber = True
End Function

' Форматирует число для диагностического журнала с точкой как разделителем.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' Утилитарный минимум без WorksheetFunction в расчетном ядре.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

' Утилитарный максимум без WorksheetFunction в расчетном ядре.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function
