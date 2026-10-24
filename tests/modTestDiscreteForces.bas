Attribute VB_Name = "modTestDiscreteForces"
Option Explicit

' ==================== ДЛЯ ТЕСТОВ ====================
' Полная прежняя сумма и групповой интеграл одной сетки, затем настоящий Newton.
' Диаграммы синтетические: проверяется алгоритм, а не нормативная калибровка.
Private Declare PtrSafe Function TestCounter Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef value As Currency) As Long
Private Declare PtrSafe Function TestFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef value As Currency) As Long
Private mTestReport As String, mTestPassed As Long, mTestFailed As Long ' Сводный отчёт без построчного вывода миллионов сумм.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет малые и целевые размеры, неоднородные площади, оба знака и узлы диаграммы.
Public Function RunDiscreteForceTests() As String
    mTestReport = "": mTestPassed = 0: mTestFailed = 0
    On Error GoTo Failed
    Dim size As Variant, section As CSectionModel, concrete As CMaterialDiagram, steel As CMaterialDiagram
    Dim frame As CSectionNumericalFrame, integrator As CDiscreteForceIntegrator
    Dim i As Long, angle As Double, eps As Double, kx As Double, ky As Double, n As Double, mx As Double, my As Double
    Dim groupedN As Double, groupedMx As Double, groupedMy As Double, grossN As Double, grossMx As Double, grossMy As Double
    Dim started As Double, buildSeconds As Double, directSeconds As Double, fastSeconds As Double, error As Double, accepted As Boolean
    For Each size In Array(50, 500, 5000, 7500, 10000, 20000, 50000)
        Set section = TestBuildForceModel(CLng(size))
        TestForceDiagrams concrete, steel
        Set frame = New CSectionNumericalFrame: frame.Initialize section
        Set integrator = New CDiscreteForceIntegrator
        started = TestForceNow(): integrator.Initialize section, frame.OriginX, frame.OriginY: buildSeconds = TestForceNow() - started
        error = 0#: directSeconds = 0#: fastSeconds = 0#
        For i = 0 To 47
            angle = (i + 0.173) * 2.39996322972865
            eps = -0.0012 + 0.0018 * Sin(i * 0.31): kx = 0.000004 * Cos(angle): ky = 0.000004 * Sin(angle)
            If i = 0 Then eps = 0#: kx = 0#: ky = 0#
            If i = 1 Then eps = -0.0003: kx = 0.0000000000000000000001: ky = 0#
            started = TestForceNow()
            TestDirectForces section, concrete, frame.OriginX, frame.OriginY, eps, kx, ky, n, mx, my, grossN, grossMx, grossMy
            directSeconds = directSeconds + TestForceNow() - started
            started = TestForceNow(): accepted = integrator.TryForces(concrete, eps, kx, ky, groupedN, groupedMx, groupedMy)
            fastSeconds = fastSeconds + TestForceNow() - started
            If accepted Then
                error = TestForceMax(error, Abs(n - groupedN) / TestForceMax(1#, grossN))
                error = TestForceMax(error, Abs(mx - groupedMx) / TestForceMax(1#, grossMx))
                error = TestForceMax(error, Abs(my - groupedMy) / TestForceMax(1#, grossMy))
            Else
                TestForceCheck "nearKnot.directFallback." & CStr(size), i = 1
            End If
        Next i
        TestForceCheck "allForces." & CStr(size), error <= 0.0000000001
        mTestReport = mTestReport & "FORCE_BENCH|count=" & CStr(size) & "|queries=48|build=" & TestForceText(buildSeconds) & _
            "|directWithErrorScale=" & TestForceText(directSeconds) & "|grouped=" & TestForceText(fastSeconds) & "|maxScaledError=" & TestForceText(error) & vbCrLf
        TestNonlinearForceSolve section, concrete, steel, 1
        TestNonlinearForceSolve section, concrete, steel, 6
        section.AddConcreteElement 700#, 200#, 100#, 1
        TestForceCheck "revisionRejected." & CStr(size), Not integrator.Matches(section.Revision, frame.OriginX, frame.OriginY)
    Next size
    GoTo Finished
Failed:
    TestForceCheck "force.runtime " & CStr(Err.Number) & ": " & Err.Description, False
Finished:
    RunDiscreteForceTests = mTestReport & "TOTAL DISCRETE FORCES: passed=" & CStr(mTestPassed) & "; failed=" & CStr(mTestFailed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Использует настоящие solve с новой и прежней оценкой проб, включая холодную подготовку.
Private Sub TestNonlinearForceSolve(ByVal section As CSectionModel, ByVal concrete As CMaterialDiagram, ByVal steel As CMaterialDiagram, ByVal caseCount As Long)
    Dim oracle As CSectionSolver: Set oracle = New CSectionSolver
    Dim loads() As Double, j As Long: ReDim loads(1 To caseCount, 1 To 3)
    For j = 1 To caseCount
        oracle.EvaluateStrainPlane section, concrete, steel, -0.0004 + (j - 1) * 0.000015, 0.000001, -0.0000006
        loads(j, 1) = oracle.Nint: loads(j, 2) = oracle.Mxint: loads(j, 3) = oracle.Myint
    Next j
    Dim enabled As Variant, solver As CSectionSolver, direct As CSectionSolver, started As Double, seconds As Double
    Dim baselines As Collection: Set baselines = New Collection
    For Each enabled In Array(False, True)
        section.TestSetForceIntegrationEnabled CBool(enabled)
        section.BeginForceIntegrationRun caseCount
        seconds = 0#
        For j = 1 To caseCount
            Set solver = New CSectionSolver
            solver.LoadSteps = 4: solver.MaxIterations = 80
            solver.ToleranceN = 0.001: solver.ToleranceMx = 0.1: solver.ToleranceMy = 0.1
            started = TestForceNow(): solver.Solve section, concrete, steel, loads(j, 1), loads(j, 2), loads(j, 3): seconds = seconds + TestForceNow() - started
            TestForceCheck "solve.converged." & CStr(section.ConcreteCount) & "." & CStr(enabled) & "." & CStr(caseCount) & "." & CStr(j), solver.Converged
            If Not CBool(enabled) Then
                baselines.Add solver
            Else
                Set direct = baselines.Item(j)
                TestForceCheck "solve.sameStatus." & CStr(section.ConcreteCount), solver.Converged = direct.Converged And solver.FailureCode = direct.FailureCode And solver.Iterations = direct.Iterations
                TestForceCheck "solve.sameEquilibrium." & CStr(section.ConcreteCount), Abs(solver.Nint - direct.Nint) <= 0.001 And Abs(solver.Mxint - direct.Mxint) <= 0.1 And Abs(solver.Myint - direct.Myint) <= 0.1
                TestForceCheck "solve.samePlane." & CStr(section.ConcreteCount), Abs(solver.Epsilon0 - direct.Epsilon0) <= 0.0000000001 And Abs(solver.KappaX - direct.KappaX) <= 0.000000000001 And Abs(solver.KappaY - direct.KappaY) <= 0.000000000001
            End If
        Next j
        mTestReport = mTestReport & "SOLVE_BENCH|count=" & CStr(section.ConcreteCount) & "|cases=" & CStr(caseCount) & "|fast=" & CStr(enabled) & "|seconds=" & TestForceText(seconds) & "|probes=" & CStr(section.TestForceProbeCount) & vbCrLf
    Next enabled
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сетка сохраняет реальные переменные площади и исходный порядок элементов.
Private Function TestBuildForceModel(ByVal count As Long) As CSectionModel
    Dim result As CSectionModel: Set result = New CSectionModel
    Dim side As Long, i As Long, ix As Long, iy As Long, width As Double
    side = -Int(-Sqr(count)): width = 1000# / side
    For i = 1 To count
        ix = (i - 1) Mod side: iy = (i - 1) \ side
        result.AddConcreteElement -500# + (ix + 0.5) * width, -500# + (iy + 0.5) * width, _
            1000000# / count * (1# + 0.2 * Sin(i)), 1, , , "Rectangle", width, width
    Next i
    result.AddRebarElement -400#, -400#, 20#, 0#, "Periodic": result.AddRebarElement 400#, -400#, 20#, 0#, "Periodic"
    result.AddRebarElement -400#, 400#, 20#, 0#, "Periodic": result.AddRebarElement 400#, 400#, 20#, 0#, "Periodic"
    Set TestBuildForceModel = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Точки заданы явно для проверки нелинейных ветвей без обращения к нормативам.
Private Sub TestForceDiagrams(ByRef concrete As CMaterialDiagram, ByRef steel As CMaterialDiagram)
    Dim e(1 To 7) As Double, s(1 To 7) As Double
    e(1) = -0.004: e(2) = -0.002: e(3) = -0.0003: e(4) = 0#: e(5) = 0.00008: e(6) = 0.00015: e(7) = 0.0003
    s(1) = -17#: s(2) = -17#: s(3) = -9#: s(4) = 0#: s(5) = 1.1: s(6) = 1.5: s(7) = 1.5
    Set concrete = New CMaterialDiagram: concrete.InitializeFromArrays e, s, 7
    Dim se(1 To 4) As Double, ss(1 To 4) As Double
    se(1) = -0.01: se(2) = -0.002: se(3) = 0.002: se(4) = 0.01
    ss(1) = -400#: ss(2) = -400#: ss(3) = 400#: ss(4) = 400#
    Set steel = New CMaterialDiagram: steel.InitializeFromArrays se, ss, 4
    concrete.Freeze: steel.Freeze ' Повторяет неизменяемые рабочие диаграммы и штатное повторное использование точной оценки.
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимый исходный проход, дополнительно считающий масштаб ошибки по модулям сил.
Private Sub TestDirectForces(ByVal section As CSectionModel, ByVal diagram As CMaterialDiagram, ByVal ox As Double, ByVal oy As Double, _
        ByVal eps As Double, ByVal kx As Double, ByVal ky As Double, ByRef n As Double, ByRef mx As Double, ByRef my As Double, _
        ByRef grossN As Double, ByRef grossMx As Double, ByRef grossMy As Double)
    Dim i As Long, x As Double, y As Double, area As Double, force As Double
    n = 0#: mx = 0#: my = 0#: grossN = 0#: grossMx = 0#: grossMy = 0#
    For i = 1 To section.ConcreteCount
        section.ConcreteElementForSolver i, x, y, area: x = x - ox: y = y - oy
        force = diagram.StressAtStrain(eps + kx * y + ky * x) * area
        n = n + force: mx = mx + force * y: my = my + force * x
        grossN = grossN + Abs(force): grossMx = grossMx + Abs(force * y): grossMy = grossMy + Abs(force * x)
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Монотонный QPC не зависит от перехода через полночь.
Private Function TestForceNow() As Double
    Dim counter As Currency, frequency As Currency
    TestCounter counter: TestFrequency frequency: TestForceNow = counter / frequency
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Машинная запись результата без локальной десятичной запятой.
Private Function TestForceText(ByVal value As Double) As String
    TestForceText = Replace$(CStr(value), ",", ".")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Положительный масштаб не превращает почти нулевой момент в относительную ошибку.
Private Function TestForceMax(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then TestForceMax = a Else TestForceMax = b
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Фиксирует отдельное условие, сохраняя неуспешный результат для разбора.
Private Sub TestForceCheck(ByVal label As String, ByVal condition As Boolean)
    If condition Then
        mTestPassed = mTestPassed + 1: mTestReport = mTestReport & "OK: " & label & vbCrLf
    Else
        mTestFailed = mTestFailed + 1: mTestReport = mTestReport & "FAIL: " & label & vbCrLf
    End If
End Sub
