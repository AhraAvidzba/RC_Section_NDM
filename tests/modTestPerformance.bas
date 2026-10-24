Attribute VB_Name = "modTestPerformance"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: неизменность результатов и время выполнения
' ==========================================================================
' Проверяет рабочие классы и диапазоны Excel без новых расчётных допусков.
' Сохраняет точные значения. Сами проверки не входят в замер времени;
' одинаковый измерительный код используется в обеих сравниваемых книгах.

#Const PERFORMANCE_CURRENT = True

Private mPassed As Long, mFailed As Long, mReport As String


#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If

Private Type TDoubleValue
    value As Double
End Type
Private Type TDoubleBytes
    value(1 To 8) As Byte
End Type

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет рост/очистку массивов и прежнее правило первой пустой ячейки.
Public Function RunPerformanceStorageTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    TestModelStorage
    TestSnapshotExtent
    GoTo Finished
Failed:
    TestCheck "storage.runtime " & CStr(Err.Number) & ": " & Err.Description, False
Finished:
    RunPerformanceStorageTests = mReport & "TOTAL PERFORMANCE STORAGE: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Измеряет действующий Add модели. Числа результата проверяются внешним A/B;
' индекс/площадь/суммы считаются после таймера, не входят в preparationSeconds.
Public Function TestMeasurePerformanceModel(ByVal count As Long) As String
    Dim model As CSectionModel, i As Long, t As Double, elapsed As Double
    Set model = New CSectionModel
    t = TestPerformanceNow()
    For i = 1 To count
        model.AddConcreteElement CDbl(i), -CDbl(i), 100#, 1, "source" & CStr(i), vbNullString, "Rectangle", 10#, 10#
    Next i
    elapsed = TestSecondsSince(t)
    Dim area As Double, x As Double, y As Double
    For i = 1 To model.ConcreteCount
        area = area + model.ConcreteArea(i)
        x = x + model.ConcreteX(i): y = y + model.ConcreteY(i)
    Next i
    TestMeasurePerformanceModel = "seconds=" & TestNumberText(elapsed) & "; count=" & CStr(model.ConcreteCount) & _
        "; area=" & TestNumberText(area) & "; x=" & TestNumberText(x) & "; y=" & TestNumberText(y) & _
        "; last=" & model.ConcreteID(count)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает прежнее поячеечное и обычное блочное чтение одного диапазона.
' Данные и размер одинаковы; временный лист удаляется также при ошибке.
Public Function TestMeasurePerformanceExtent(ByVal optimized As Boolean, ByVal count As Long) As String
    On Error GoTo Failed
    Dim ws As Object, values() As Variant, i As Long, t As Double, elapsed As Double, found As Long
    Set ws = ThisWorkbook.Worksheets.Add
    ReDim values(1 To count, 1 To 1)
    For i = 1 To count: values(i, 1) = "row" & CStr(i): Next i
    ws.Range("A1").Resize(count, 1).Value2 = values
    t = TestPerformanceNow()
    If optimized Then
        found = CLng(Application.Run("'" & ThisWorkbook.Name & "'!modWorkbookCalculation.SnapshotAnchoredExtent", ws.Range("A1"), False))
    Else
        found = TestOldExtent(ws.Range("A1"), False, 0)
    End If
    elapsed = TestSecondsSince(t)
    TestMeasurePerformanceExtent = "seconds=" & TestNumberText(elapsed) & "; count=" & CStr(found)
    ws.Delete
    Exit Function
Failed:
    Dim errorNumber As Long, description As String
    errorNumber = Err.Number: description = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo 0
    Err.Raise errorNumber, "TestMeasurePerformanceExtent", description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' поля по обе стороны нескольких границ резерва и Clear.
Private Sub TestModelStorage()
    Dim model As CSectionModel, i As Long, revision As Long, value As Double, errorNumber As Long
    Set model = New CSectionModel
    For i = 1 To 257
        model.AddConcreteElement CDbl(i) / 8#, -CDbl(i) / 16#, CDbl(i) + 1#, 2, "source" & CStr(i), "handle" & CStr(i)
        model.AddRebarElement CDbl(i) / 32#, CDbl(i) / 64#, 20#, CDbl(i) + 2#, "Periodic", 3, "bar" & CStr(i), "r" & CStr(i)
    Next i
    TestCheck "model.count", model.ConcreteCount = 257 And model.RebarCount = 257
    For i = 1 To 257
        TestCheck "model.concrete." & CStr(i), model.ConcreteID(i) = "C" & CStr(i) And _
            model.ConcreteX(i) = CDbl(i) / 8# And model.ConcreteY(i) = -CDbl(i) / 16# And _
            model.ConcreteArea(i) = CDbl(i) + 1# And model.ConcreteSourceHandle(i) = "handle" & CStr(i)
        TestCheck "model.rebar." & CStr(i), model.RebarID(i) = "R" & CStr(i) And _
            model.RebarX(i) = CDbl(i) / 32# And model.RebarY(i) = CDbl(i) / 64# And _
            model.RebarArea(i) = CDbl(i) + 2# And model.RebarSourceName(i) = "bar" & CStr(i)
    Next i
    revision = model.Revision
    On Error Resume Next
    model.AddConcreteElement 0#, 0#, 0#: errorNumber = Err.Number: Err.Clear
    On Error GoTo 0
    TestCheck "model.invalidConcrete", errorNumber <> 0 And model.ConcreteCount = 257 And model.Revision = revision
    On Error Resume Next
    model.AddRebarElement 0#, 0#, 0#, 0#, "Periodic": errorNumber = Err.Number: Err.Clear
    On Error GoTo 0
    TestCheck "model.invalidRebar", errorNumber <> 0 And model.RebarCount = 257 And model.Revision = revision
    On Error Resume Next
    value = model.ConcreteX(258): errorNumber = Err.Number: Err.Clear
    On Error GoTo 0
    TestCheck "model.tailHidden", errorNumber <> 0
    model.Clear
    TestCheck "model.clear", model.ConcreteCount = 0 And model.RebarCount = 0 And model.Revision = revision + 1
    model.AddConcreteElement 1#, 2#, 3#: model.AddRebarElement 4#, 5#, 6#, 7#, "Smooth"
    TestCheck "model.rebuild", model.ConcreteID(1) = "C1" And model.ConcreteX(1) = 1# And model.RebarID(1) = "R1" And model.RebarArea(1) = 7#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' чтение одной ячейки и массива, пустых значений, формул, пробелов, ошибок и края
' листа.
Private Sub TestSnapshotExtent()
    On Error GoTo Failed
    Dim ws As Object, anchor As Object, values() As Variant, i As Long
    Set ws = ThisWorkbook.Worksheets.Add
    ReDim values(1 To 2050, 1 To 1)
    For i = 1 To 2050: values(i, 1) = "r" & CStr(i): Next i
    ws.Range("B2").Resize(2050, 1).Value2 = values
    TestCheckExtent ws.Range("B2"), False, 0, "rows.blocks"
    ws.Range("B1027").Value2 = "   "
    TestCheckExtent ws.Range("B2"), False, 0, "rows.space"
    ws.Range("B1027").Formula = "="""""
    TestCheckExtent ws.Range("B2"), False, 0, "rows.formulaEmpty"
    ws.Range("B1027").Value2 = CVErr(2042)
    TestCheckExtent ws.Range("B2"), False, 0, "rows.errorNotBlank"
    ws.Range("B1027").ClearContents
    TestCheckExtent ws.Range("B2"), False, 0, "rows.empty"
    TestCheckExtent ws.Range("C2"), False, 0, "rows.blank"
    ws.Range("D4:F4").Value2 = "header"
    TestCheckExtent ws.Range("D4"), True, 256, "columns"
    TestCheckExtent ws.Range("D4"), True, 1, "columns.limitOne"
    Set anchor = ws.Cells(ws.Rows.Count, ws.Columns.Count)
    TestCheckExtent anchor, False, 0, "edge.emptyRow"
    TestCheckExtent anchor, True, 256, "edge.emptyColumn"
    anchor.Value2 = "last"
    TestCheckExtent anchor, False, 0, "edge.scalarRow"
    TestCheckExtent anchor, True, 256, "edge.scalarColumn"
    ws.Delete
    Exit Sub
Failed:
    Dim errorNumber As Long, description As String
    errorNumber = Err.Number: description = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo 0
    Err.Raise errorNumber, "TestSnapshotExtent", description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает блочное определение занятой области с прежним поклеточным эталоном при одинаковом ограничении поиска.
Private Sub TestCheckExtent(ByVal anchor As Object, ByVal columns As Boolean, ByVal limit As Long, ByVal label As String)
    Dim actual As Long, expected As Long
    expected = TestOldExtent(anchor, columns, limit)
    actual = CLng(Application.Run("'" & ThisWorkbook.Name & "'!modWorkbookCalculation.SnapshotAnchoredExtent", anchor, columns, limit))
    TestCheck label, actual = expected
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимый прежний oracle: именно первая Trim/CStr-пустая, не UsedRange.
Private Function TestOldExtent(ByVal anchor As Object, ByVal columns As Boolean, ByVal limit As Long) As Long
    Dim available As Long, i As Long, value As Variant
    If columns Then available = anchor.Worksheet.Columns.Count - anchor.Column + 1 Else available = anchor.Worksheet.Rows.Count - anchor.Row + 1
    If limit > 0 And limit < available Then available = limit
    For i = 0 To available - 1
        If columns Then value = anchor.Offset(0, i).Value2 Else value = anchor.Offset(i, 0).Value2
        If Len(Trim$(CStr(value))) = 0 Then Exit Function
        TestOldExtent = TestOldExtent + 1
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' фиксирует только неуспешные assertions, не меняет expected/tolerance.
Private Sub TestCheck(ByVal label As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & label & vbCrLf
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' длительность по одному монотонному таймеру, не VBA Timer.
Private Function TestSecondsSince(ByVal start As Double) As Double
    TestSecondsSince = TestPerformanceNow() - start
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' монотонный высокоточный таймер; одинаковый код у A и B.
Private Function TestPerformanceNow() As Double
    Dim counter As Currency, frequency As Currency
    If QueryPerformanceCounter(counter) = 0 Or QueryPerformanceFrequency(frequency) = 0 Then _
        Err.Raise vbObjectError + 4980, "modTestPerformance", "Высокоточный таймер недоступен."
    TestPerformanceNow = counter / frequency
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сериализует все биты Double, не округляет и не меняет tolerance.
Private Function TestDoubleBits(ByVal value As Double) As String
    Dim number As TDoubleValue, bytes As TDoubleBytes, i As Long
    number.value = value: LSet bytes = number
    For i = 1 To 8: TestDoubleBits = TestDoubleBits & Right$("0" & Hex$(bytes.value(i)), 2): Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет инженерные величины решателя без служебных счётчиков.
' Матрица учитывается только при проверке полного вычисления НДС или обычного Solve.
Private Function TestSolverFingerprint(ByVal solver As CSectionSolver, Optional ByVal includeK As Boolean = True) As String
    Dim values As Variant, value As Variant, i As Long, j As Long
    values = Array(solver.Converged, CLng(solver.FailureCode), solver.StopReason, solver.Epsilon0, solver.KappaX, solver.KappaY, _
        solver.Nint, solver.Mxint, solver.Myint, solver.ResidualN, solver.ResidualMx, solver.ResidualMy, _
        solver.RelativeResidualN, solver.RelativeResidualMx, solver.RelativeResidualMy, _
        solver.MinConcreteStrain, solver.MaxConcreteStrain, solver.MaxConcreteCompressionStress, _
        solver.MinConcreteStress, solver.MaxConcreteStress, solver.HasConcreteCompression, solver.HasConcreteTension, _
        solver.MinSteelStrain, solver.MaxSteelStrain, solver.MaxAbsSteelStrain, solver.MinSteelStress, solver.MaxSteelStress, _
        solver.HasSteelCompression, solver.HasSteelTension, solver.Iterations, solver.TestLoadStepsCompleted, _
        solver.LastResidualNorm, solver.LastStepNorm, solver.LastMatrixConditionEstimate, solver.TestMatrixRestartCount)
    For Each value In values
        If VarType(value) = vbDouble Then
            TestSolverFingerprint = TestSolverFingerprint & "|" & TestDoubleBits(CDbl(value))
        Else
            TestSolverFingerprint = TestSolverFingerprint & "|" & CStr(value)
        End If
    Next value
    If includeK Then
        For i = 1 To 3: For j = 1 To 3
            TestSolverFingerprint = TestSolverFingerprint & "|" & TestDoubleBits(solver.Tangent(i, j))
        Next j: Next i
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' штатные диаграммы/узлы, две численные ветви и границы одного Solve.
' Frozen baseline не компилирует новые Friend API; численный oracle не меняется.
Public Function RunPerformanceNumericTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
#If PERFORMANCE_CURRENT Then
    Dim section As CSectionModel, concrete As CMaterialDiagram, steel As CMaterialDiagram
    Dim frozenConcrete As CMaterialDiagram, frozenSteel As CMaterialDiagram
    Set section = TestNumericSection(8)
    TestNumericMaterials concrete, steel
    Set frozenConcrete = concrete.Clone: Set frozenSteel = steel.Clone
    Dim diagram As Variant, i As Long, shift As Variant, eps As Double
    Dim full As CSectionSolver, physical As CSectionSolver
    Set full = New CSectionSolver: Set physical = New CSectionSolver
    For Each diagram In Array(frozenConcrete, frozenSteel)
        For i = 1 To diagram.PointCount
            For Each shift In Array(-0.000000000001, 0#, 0.000000000001)
                eps = diagram.PointStrain(i) + CDbl(shift)
                full.EvaluateStrainPlane section, frozenConcrete, frozenSteel, eps, 0#, 0#
                physical.EvaluatePhysicalStrainPlane section, frozenConcrete, frozenSteel, eps, 0#, 0#
                TestCheck "physical.exact." & CStr(i) & "." & CStr(shift), TestSolverFingerprint(full, False) = TestSolverFingerprint(physical, False)
                TestCheck "physical.noK", physical.TestTangentBuildCount = 0 And full.TestTangentBuildCount = 1
            Next shift
        Next i
    Next diagram
    Dim method As Variant, n As Double, mx As Double, my As Double, hits As Long
    Dim plain As CSectionSolver, cached As CSectionSolver
    For Each method In Array("Newton", "Secant")
        Set plain = New CSectionSolver: Set cached = New CSectionSolver
        plain.SolverMethod = CStr(method): cached.SolverMethod = CStr(method)
        plain.DiagnosticsEnabled = False: cached.DiagnosticsEnabled = False
        For i = 1 To 30
            n = (-1#) ^ i * (100000# + 11000# * i)
            mx = (i - 15#) * 1000000#: my = ((7 * i Mod 31) - 15#) * 700000#
            plain.Solve section, concrete, steel, n, mx, my
            cached.Solve section, frozenConcrete, frozenSteel, n, mx, my
            TestCheck "solve.exact." & CStr(method) & "." & CStr(i), TestSolverFingerprint(plain) = TestSolverFingerprint(cached)
            TestCheck "solve.mutableNoReuse", plain.TestEvaluationReuseCount = 0
            hits = hits + cached.TestEvaluationReuseCount
        Next i
        TestCheck "solve.reuse." & CStr(method), hits > 0
    Next method
    section.AddConcreteElement 205#, 0#, 100#, 1, "mutation"
    plain.Solve section, concrete, steel, -100000#, 2000000#, -3000000#
    cached.Solve section, frozenConcrete, frozenSteel, -100000#, 2000000#, -3000000#
    TestCheck "solve.changedModel", TestSolverFingerprint(plain) = TestSolverFingerprint(cached)
    cached.ClearResult
    TestCheck "solve.clear", cached.TestEvaluationReuseCount = 0 And cached.TestTangentBuildCount = 0
#Else
    TestCheck "numeric.currentOnly", False
#End If
    GoTo Finished
Failed:
    TestCheck "numeric.runtime " & CStr(Err.Number) & ": " & Err.Description, False
Finished:
    RunPerformanceNumericTests = mReport & "TOTAL PERFORMANCE NUMERIC: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' детерминированная сетка и четыре стержня; без Excel в ядре.
Private Function TestNumericSection(ByVal divisions As Long) As CSectionModel
    Dim section As CSectionModel, i As Long, j As Long, step As Double
    Set section = New CSectionModel: step = 400# / divisions
    For i = 1 To divisions: For j = 1 To divisions
        section.AddConcreteElement -200# + (i - 0.5) * step, -200# + (j - 0.5) * step, step * step, _
            1, "cell", vbNullString, "Rectangle", step, step
    Next j: Next i
    For i = -1 To 1 Step 2: For j = -1 To 1 Step 2
        section.AddRebarElement 140# * i, 140# * j, 20#, 0#, "Periodic", 1, "bar"
    Next j: Next i
    Set TestNumericSection = section
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' одинаковые изменяемые и защищённые от изменений копии
' с явно заданными материалами без значений по умолчанию из Config.
Private Sub TestNumericMaterials(ByRef concrete As CMaterialDiagram, ByRef steel As CMaterialDiagram)
    Dim eps(1 To 5) As Double, stress(1 To 5) As Double
    eps(1) = -0.0035: eps(2) = -0.002: eps(3) = 0#: eps(4) = 0.00015: eps(5) = 0.0003
    stress(1) = -30#: stress(2) = -30#: stress(3) = 0#: stress(4) = 2#: stress(5) = 2#
    Set concrete = New CMaterialDiagram: concrete.InitializeFromArrays eps, stress, 5
    eps(1) = -0.02: eps(2) = -0.002: eps(3) = 0#: eps(4) = 0.002: eps(5) = 0.02
    stress(1) = -400#: stress(2) = -400#: stress(3) = 0#: stress(4) = 400#: stress(5) = 400#
    Set steel = New CMaterialDiagram: steel.InitializeFromArrays eps, stress, 5
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' одинаковые 1/10/30 независимых задач; preparation/core отдельно.
' Fingerprint и assertions выполняются после таймера ядра.
Public Function TestMeasurePerformanceSolver(ByVal method As String, ByVal divisions As Long, ByVal count As Long) As String
    Dim start As Double, prepared As Double, elapsed As Double, section As CSectionModel
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, solvers() As CSectionSolver, i As Long
    start = TestPerformanceNow()
    Set section = TestNumericSection(divisions): TestNumericMaterials concrete, steel
    Set concrete = concrete.Clone: Set steel = steel.Clone
    ReDim solvers(1 To count)
    prepared = TestSecondsSince(start)
    start = TestPerformanceNow()
    For i = 1 To count
        Set solvers(i) = New CSectionSolver
        solvers(i).SolverMethod = method: solvers(i).DiagnosticsEnabled = False
        solvers(i).Solve section, concrete, steel, (-1#) ^ i * (100000# + 11000# * i), _
            (i - 15#) * 1000000#, ((7 * i Mod 31) - 15#) * 700000#
    Next i
    elapsed = TestSecondsSince(start)
    Dim fingerprint As String, evaluations As Long, builds As Long, reuse As Long
    For i = 1 To count
        fingerprint = fingerprint & vbLf & CStr(i) & TestSolverFingerprint(solvers(i))
        evaluations = evaluations + solvers(i).TestInternalForceEvaluationCount
#If PERFORMANCE_CURRENT Then
        builds = builds + solvers(i).TestTangentBuildCount: reuse = reuse + solvers(i).TestEvaluationReuseCount
#End If
    Next i
    TestMeasurePerformanceSolver = "seconds=" & TestNumberText(elapsed) & Chr$(30) & "preparation=" & TestNumberText(prepared) & _
        Chr$(30) & "elements=" & CStr(section.ConcreteCount) & Chr$(30) & "bars=" & CStr(section.RebarCount) & _
        Chr$(30) & "evaluations=" & CStr(evaluations) & Chr$(30) & "K=" & CStr(builds) & _
        Chr$(30) & "reuse=" & CStr(reuse) & Chr$(30) & "fingerprint=" & fingerprint
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' побитовый снимок механических A/I, не времени вычисления.
Private Function TestPropertiesFingerprint(ByVal props As CSectionPropertiesCalculator) As String
    Dim value As Variant
    For Each value In Array(props.Area, props.StaticMomentX, props.StaticMomentY, props.CentroidX, props.CentroidY, _
            props.Ix, props.Iy, props.Ixy, props.Ixc, props.Iyc, props.Ixyc, props.PrincipalI1, props.PrincipalI2, _
            props.PrincipalAngleRad, props.RadiusX, props.RadiusY, props.PrincipalRadius1, props.PrincipalRadius2)
        TestPropertiesFingerprint = TestPropertiesFingerprint & "|" & TestDoubleBits(CDbl(value))
    Next value
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' точный порядок/координаты/дуги и площадь готовой области.
Private Function TestRegionFingerprint(ByVal region As CConcreteRegion) As String
    If region Is Nothing Then TestRegionFingerprint = "Nothing": Exit Function
    Dim edges() As TRegionEdge, count As Long, i As Long
    region.CopyEdges edges, count
    TestRegionFingerprint = region.Source & "|" & CStr(region.SectionRevision) & "|" & CStr(region.LoopCount) & "|" & TestDoubleBits(region.Area)
    For i = 1 To count
        TestRegionFingerprint = TestRegionFingerprint & "|" & CStr(edges(i).LoopID) & TestDoubleBits(edges(i).X1) & TestDoubleBits(edges(i).Y1) & _
            TestDoubleBits(edges(i).X2) & TestDoubleBits(edges(i).Y2) & TestDoubleBits(edges(i).CenterX) & TestDoubleBits(edges(i).CenterY) & TestDoubleBits(edges(i).Sweep)
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' общие геометрические данные не смешиваются между сочетаниями
' и не удерживаются готовым результатом СП 35. Изменение модели, контуров
' или модулей материалов требует новой подготовки.
Public Function RunPerformancePreparationTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
#If PERFORMANCE_CURRENT Then
    Dim section As CSectionModel, query As CSectionGeometryQuery, fresh As CSectionGeometryQuery
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, first As CConcreteRegion, second As CConcreteRegion
    Set section = TestNumericSection(8): TestNumericMaterials concrete, steel
    Set query = New CSectionGeometryQuery: query.Initialize section
    Dim i As Long, j As Long, sharedData As CSP35CrackData, freshData As CSP35CrackData, members As Variant
    For i = 1 To 30
        Set fresh = New CSectionGeometryQuery: fresh.Initialize section
        Set first = query.ConcreteDomain: Set second = fresh.ConcreteDomain
        TestCheck "domain.exact." & CStr(i), TestRegionFingerprint(first) = TestRegionFingerprint(second)
        first.SetOwner "changed clone", "SP35", "G1"
        Set second = query.ConcreteDomain
        TestCheck "domain.clone", Len(second.OwnerID) = 0
        Set sharedData = New CSP35CrackData: Set freshData = New CSP35CrackData
        sharedData.Prepare section, 0.0004 + i * 0.000001, 0.0000001, -0.00000015, steel, 10#, 50#, "Max", "6d", False, 0.2, query
        freshData.Prepare section, 0.0004 + i * 0.000001, 0.0000001, -0.00000015, steel, 10#, 50#, "Max", "6d", False
        TestCheck "sp35.count", sharedData.CandidateCount = freshData.CandidateCount And sharedData.GroupCount = freshData.GroupCount
        For j = 1 To sharedData.CandidateCount
            TestCheck "sp35.region", TestRegionFingerprint(sharedData.CandidateRegion(j)) = TestRegionFingerprint(freshData.CandidateRegion(j))
            TestCheck "sp35.beta", TestDoubleBits(sharedData.CandidateBetaDiameterSum(j)) = TestDoubleBits(freshData.CandidateBetaDiameterSum(j))
            TestCheck "sp35.stress", TestDoubleBits(sharedData.CandidateSigmaS(j)) = TestDoubleBits(freshData.CandidateSigmaS(j))
        Next j
        sharedData.Freeze
        members = sharedData.GroupMembers(1)
        TestCheck "sp35.published", members(1) > 0
    Next i
    TestCheck "domain.singleBuild", query.TestDomainBuildCount = 1
    section.Contours.AddContourCircle "opening", 0#, 0#, 30#, "test", "opening", "Opening"
    query.EnsureSection section
    Set fresh = New CSectionGeometryQuery: fresh.Initialize section
    Set first = query.ConcreteDomain: Set second = fresh.ConcreteDomain
    TestCheck "domain.contoursOnly", TestRegionFingerprint(first) = TestRegionFingerprint(second) And query.TestDomainBuildCount = 2
    Set section = TestNumericSection(8)
    query.EnsureSection section
    Set fresh = New CSectionGeometryQuery: fresh.Initialize section
    Set first = query.ConcreteDomain: Set second = fresh.ConcreteDomain
    TestCheck "domain.identity", TestRegionFingerprint(first) = TestRegionFingerprint(second) And query.TestDomainBuildCount = 3
    Dim props As CSectionPropertiesCalculator, oracle As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator: Set oracle = New CSectionPropertiesCalculator
    Dim nx As Double, ny As Double, low As Double, high As Double, oldLow As Double, oldHigh As Double
    For i = 1 To 30
        props.EnsureTransformed section, 30000#, 200000#: oracle.CalculateTransformedByModuli section, 30000#, 200000#
        TestCheck "properties.exact", TestPropertiesFingerprint(props) = TestPropertiesFingerprint(oracle)
        nx = Cos(CDbl(i)): ny = Sin(CDbl(i))
        For j = 1 To 2
            props.CalculateProjection section, nx, ny, False, low, high
            oracle.CalculateProjection section, nx, ny, False, oldLow, oldHigh
            TestCheck "properties.projection", TestDoubleBits(low) = TestDoubleBits(oldLow) And TestDoubleBits(high) = TestDoubleBits(oldHigh)
            TestCheck "properties.concreteI", TestDoubleBits(props.ConcreteInertiaAboutAxis(section, nx, ny, 3#, -2#)) = _
                TestDoubleBits(oracle.ConcreteInertiaAboutAxis(section, nx, ny, 3#, -2#))
            TestCheck "properties.rebarI", TestDoubleBits(props.RebarInertiaAboutAxis(section, nx, ny, 3#, -2#)) = _
                TestDoubleBits(oracle.RebarInertiaAboutAxis(section, nx, ny, 3#, -2#))
        Next j
    Next i
    TestCheck "properties.singleBuild", props.TestPreparationBuildCount = 1
    props.EnsureTransformed section, 31000#, 200001#: oracle.CalculateTransformedByModuli section, 31000#, 200001#
    TestCheck "properties.changedModuli", TestPropertiesFingerprint(props) = TestPropertiesFingerprint(oracle) And props.TestPreparationBuildCount = 2
    section.AddConcreteElement 210#, 0#, 100#, 1, "mutation"
    props.EnsureTransformed section, 31000#, 200001#: oracle.CalculateTransformedByModuli section, 31000#, 200001#
    TestCheck "properties.changedModel", TestPropertiesFingerprint(props) = TestPropertiesFingerprint(oracle) And props.TestPreparationBuildCount = 3
    props.EnsureConcrete section: oracle.CalculateConcrete section
    TestCheck "properties.changedRepresentation", TestPropertiesFingerprint(props) = TestPropertiesFingerprint(oracle) And props.TestPreparationBuildCount = 4
    props.Clear: props.EnsureConcrete section
    TestCheck "properties.clear", TestPropertiesFingerprint(props) = TestPropertiesFingerprint(oracle) And props.TestPreparationBuildCount = 5
#Else
    TestCheck "preparation.currentOnly", False
#End If
    GoTo Finished
Failed:
    TestCheck "preparation.runtime " & CStr(Err.Number) & ": " & Err.Description, False
Finished:
    RunPerformancePreparationTests = mReport & "TOTAL PERFORMANCE PREPARATION: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' машинная запись времени с точкой для одинакового A/B parser.
Private Function TestNumberText(ByVal value As Double) As String
    TestNumberText = Replace$(Format$(value, "0.###############"), ",", ".")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' только частная A/B-копия. 1/10 LC - префиксы 30 разных LC.
' Порядок меняет позиции, но не ID/нагрузки. Адреса читаются по именам/ключам.
Public Sub TestConfigurePerformanceFixture(ByVal family As String, ByVal count As Long, _
        ByVal method As String, Optional ByVal shape As String = "Saved", _
        Optional ByVal order As String = "Normal", Optional ByVal extension As String = "Yes")
    If count < 1 Or count > 30 Then Err.Raise 5, "TestConfigurePerformanceFixture", "Test count must be 1..30."
    TestPerformanceSetting "Solver.Method", method
    TestPerformanceSetting "General.DiagramExtension", extension
    TestPerformanceSetting "Calculation.ZeroMomentPerDepth", "0"
    If shape <> "Saved" Then
        TestPerformanceSetting "Geometry.Source", "Generated"
        TestPerformanceSetting "Geometry.Type", shape
    End If
    Dim capacity As Boolean, crack As Boolean, stability As Boolean
    capacity = InStr(family, "CAPACITY") > 0 Or family = "MIXED"
    crack = InStr(family, "CRACK") > 0 Or family = "MIXED"
    stability = InStr(family, "STABILITY") > 0 Or family = "MIXED"
    If InStr(family, "SP63") > 0 Then
        If crack Then TestPerformanceSetting "SLS.Crack.Code", "SP63"
        If stability Then TestPerformanceSetting "Stability.Code", "SP63"
    Else
        If crack Then TestPerformanceSetting "SLS.Crack.Code", "SP35"
        If stability Then TestPerformanceSetting "Stability.Code", "SP35"
    End If
    If capacity Then
        If InStr(family, "MULTIPLIER") > 0 Then
            TestPerformanceSetting "Capacity.SolutionStrategy", "LoadMultiplier"
        Else
            TestPerformanceSetting "Capacity.SolutionStrategy", "Auto"
        End If
    End If
    Dim profiles As Object, data As Variant, row As Long, col As Long, key As String, headerRow As Long
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    data = profiles.Value2
    For row = 1 To UBound(data, 1)
        If CStr(data(row, 3)) = "PR1" Then headerRow = row: Exit For
    Next row
    If headerRow = 0 Then Err.Raise 5, "TestConfigurePerformanceFixture", "Missing profile header."
    For row = 1 To UBound(data, 1)
        key = CStr(data(row, 2))
        For col = 3 To UBound(data, 2)
            If Left$(CStr(data(headerRow, col)), 2) <> "PR" Then GoTo NextProfileColumn
            Select Case key
                Case "Calculation.Strength.DirectState": data(row, col) = "Yes"
                Case "Calculation.Strength.Capacity": data(row, col) = IIf(capacity, "Yes", "No")
                Case "Calculation.Crack.Width": data(row, col) = IIf(crack, "Yes", "No")
                Case "Calculation.Stability.Enabled": data(row, col) = IIf(stability And col = 3, "Yes", "No")
            End Select
NextProfileColumn:
        Next col
    Next row
    profiles.Value2 = data
    Dim settings As CSystemSettingsReader, units As CUnitSystem
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Dim loads As Object, duration As Object, values() As Variant, durations() As Variant
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    If loads.Rows.Count < 31 Then Err.Raise 5, "TestConfigurePerformanceFixture", "Private fixture needs 30 load rows."
    ReDim values(1 To loads.Rows.Count - 1, 1 To loads.Columns.Count)
    Set duration = ThisWorkbook.Names.Item("rngStabilityDurationLoads").RefersToRange
    ReDim durations(1 To duration.Rows.Count - 1, 1 To duration.Columns.Count)
    Dim i As Long, j As Long, n As Double, mx As Double, my As Double, paths As Variant
    paths = Array("Auto", "lambda*Mx", "lambda*My", "lambda*Mxy", "lambda*N", "lambda*NMxy")
    For i = 1 To count
        j = i
        If order = "Reverse" Then j = count + 1 - i
        If order = "Shuffle" Then j = ((i - 1) * 7 Mod count) + 1
        n = (-1#) ^ j * (100000# + 11000# * j)
        mx = (j - 15#) * 1000000#: my = ((7 * j Mod 31) - 15#) * 700000#
        If crack Then
            n = -50000# + (j Mod 4) * 20000#
            mx = (15# + j) * 10000000#: my = ((j Mod 5) - 2#) * 30000000#
            If j Mod 7 = 0 Then n = -100000#: mx = 0#: my = 0#
        End If
        If stability Then n = -100000# - j * 15000#: mx = j * 1000000#: my = (j Mod 6) * 800000#
        If j = 1 And Not crack Then n = -111000#: mx = 0#: my = 0#
        If family = "MIXED" And j = 15 Then n = -1000000000#: mx = 10000000000#
        If family = "TINY" Then n = -j * 0.00000001: mx = j * 0.00000002: my = -j * 0.00000003
        If family = "REPEAT" Then n = -120000#: mx = 15000000#: my = -6000000#
        values(i, 1) = "PERF_" & Format$(j, "00")
        values(i, 2) = n / units.InputForceToInternal(1#)
        values(i, 3) = mx / units.InputMomentMxToInternal(1#)
        values(i, 4) = my / units.InputMomentMyToInternal(1#)
        values(i, 5) = "PR1"
        If stability And j Mod 5 = 0 Then values(i, 5) = "PR2"
        values(i, 6) = paths((j - 1) Mod 6)
        values(i, 7) = "Различное сочетание " & CStr(j)
        If family = "MIXED" And j = 8 Then values(i, 2) = "invalid"
        If i <= UBound(durations, 1) Then
            durations(i, 1) = values(i, 1)
            durations(i, 2) = n * 0.5 / units.InputForceToInternal(1#)
            durations(i, 3) = mx * 0.5 / units.InputMomentMxToInternal(1#)
            durations(i, 4) = my * 0.5 / units.InputMomentMyToInternal(1#)
        End If
    Next i
    loads.Offset(1, 0).Resize(UBound(values, 1), UBound(values, 2)).Value2 = values
    duration.Offset(1, 0).Resize(UBound(durations, 1), UBound(durations, 2)).Value2 = durations
    If Left$(family, 9) = "IMPORTED_" Then TestPreparePerformanceImportedSnapshot units
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохранение и чтение импортированной геометрии с точными дугами
' и двумя отверстиями без запуска AutoCAD. В книгах A и B одинаковы сетка,
' контуры и стержни; точный контур не заменяет расчётную сетку.
Private Sub TestPreparePerformanceImportedSnapshot(ByVal units As CUnitSystem)
    Dim section As CSectionModel, query As CSectionGeometryQuery, region As CConcreteRegion
    Dim props As CSectionPropertiesCalculator, writer As CNDMResultsWriter
    Dim x As Double, y As Double, i As Long, angle As Double, pi As Double
    Set section = New CSectionModel: section.SourceType = "AutoCADImport"
    Set query = New CSectionGeometryQuery
    Set region = query.CircleRegion(0#, 0#, 200#)
    Set region = query.SubtractRegion(region, query.CircleRegion(-45#, 60#, 25#))
    Set region = query.SubtractRegion(region, query.CircleRegion(45#, 60#, 25#))
    For x = -195# To 195# Step 10#
        For y = -195# To 195# Step 10#
            If query.ContainsPoint(region, x, y) Then
                section.AddConcreteElement x, y, 100#, 1, "CAD_CELL_" & CStr(section.ConcreteCount + 1), _
                    vbNullString, "Rectangle", 10#, 10#
            End If
        Next y
    Next x
    pi = 4# * Atn(1#)
    For i = 1 To 8
        angle = (i - 1) * pi / 4#
        section.AddRebarElement 170# * Cos(angle), 170# * Sin(angle), 30#, pi * 30# ^ 2 / 4#, "Periodic", _
            1, "CAD_BAR_" & CStr(i)
    Next i
    section.Contours.AddMaterialRegion region, "PERFORMANCE_CAD", "Точный контур тестового снимка AutoCAD."
    Set props = New CSectionPropertiesCalculator: props.PrepareSnapshot section
    Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, section, props, units
    TestPerformanceSetting "Geometry.Source", "AutoCAD"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' настройка частной копии по фактическому ключу.
Private Sub TestPerformanceSetting(ByVal key As String, ByVal value As String)
    Dim table As Object, data As Variant, row As Long
    Set table = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    data = table.Value2
    For row = 1 To UBound(data, 1)
        If CStr(data(row, 1)) = key Then table.Cells(row, 2).Value2 = value: Exit Sub
    Next row
    Err.Raise 5, "PerformanceSetting", "Missing setting: " & key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Full - настоящий макрос; Staged - те же владельцы с QPC.
' Assertions/fingerprint вне таймера; ни один объект не живет между вызовами.
Public Function TestMeasurePerformanceBatch(ByVal mode As String) As String
    Dim t As Double, total As Double, preparation As Double, core As Double, clear As Double
    Dim summary As Double, snapshot As Double, plot As Double, packing As Double
    Dim guard As CExcelAppStateGuard, message As String, tech As String, fingerprint As String
    On Error GoTo Failed
    t = TestPerformanceNow(): total = t
    If mode = "Full" Then
        message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
        total = TestSecondsSince(total)
    Else
        Set guard = New CExcelAppStateGuard: guard.Enter Application
        Dim settings As CSystemSettingsReader, units As CUnitSystem, section As CSectionModel
        Dim provider As CMaterialModelProvider, profiles As CCalculationProfileCatalog
        Dim batch As CBatchSectionCalculator, reader As CLoadCombinationReader, report As CExecutionReport
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
        Set provider = New CMaterialModelProvider: provider.Initialize settings, units
        Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
        Set batch = New CBatchSectionCalculator: batch.Initialize section, provider
        Set batch.ProfileCatalog = profiles: batch.ApplySettings settings, units
        Set report = New CExecutionReport: report.Initialize ThisWorkbook, settings
        Set batch.ExecutionReport = report
        Set reader = New CLoadCombinationReader: reader.LoadFromWorkbook ThisWorkbook, batch, units
        LoadStabilityDurationLoadsFromWorkbook ThisWorkbook, batch, units
        Dim table As Variant, inputError As String
        table = ReadSP35Table721FromWorkbook(ThisWorkbook, inputError): batch.SetSP35Table721 table, inputError
        Audit03ApplyLoadReferenceForTests section, settings, units, batch
        preparation = TestSecondsSince(t)
        t = TestPerformanceNow(): ClearSectionResultsForWorkbook ThisWorkbook: clear = TestSecondsSince(t)
        t = TestPerformanceNow(): batch.Execute: core = TestSecondsSince(t)
        Dim summaryWriter As CBatchResultWriter, ndm As CNDMResultsWriter, props As CSectionPropertiesCalculator
        Set summaryWriter = New CBatchResultWriter: Set ndm = New CNDMResultsWriter
        t = TestPerformanceNow(): Set props = PrepareSectionSnapshot(section, provider): packing = TestSecondsSince(t)
        t = TestPerformanceNow(): summaryWriter.WriteSummary ThisWorkbook, batch, units, section: summary = TestSecondsSince(t)
        t = TestPerformanceNow(): ndm.WriteResults ThisWorkbook, section, props, batch, units: snapshot = TestSecondsSince(t)
        t = TestPerformanceNow()
        If settings.GetRequiredBoolean("Plot.Enabled") And settings.GetRequiredBoolean("Plot.AutoUpdateAfterCalculation") Then
            If batch.StateAvailableCount > 0 Then UpdateSectionPlotForWorkbook ThisWorkbook, False
        End If
        plot = TestSecondsSince(t): total = TestSecondsSince(total)
        tech = "elements=" & CStr(section.ConcreteCount) & Chr$(30) & "bars=" & CStr(section.RebarCount) & Chr$(30) & _
            "solves=" & CStr(batch.SolverCallCount)
#If PERFORMANCE_CURRENT Then
        tech = tech & Chr$(30) & "domainBuilds=" & CStr(batch.TestCrackDomainBuildCount) & Chr$(30) & _
            "mechanicalBuilds=" & CStr(batch.TestMechanicalPreparationBuildCount) & Chr$(30) & "released=" & CStr(batch.TestRunPreparationReleased)
#End If
        Dim i As Long, result As CCombinationResult
        For i = 1 To batch.Count
            Set result = batch.ResultAt(i)
            fingerprint = fingerprint & batch.CombinationID(i) & "|" & TestMetaFingerprint(result.OverallMeta) & "|" & _
                TestMetaFingerprint(result.DirectStateMeta) & "|" & TestMetaFingerprint(result.CapacityMeta) & "|" & _
                TestMetaFingerprint(result.CrackFormationMeta) & "|" & TestMetaFingerprint(result.CrackCurrentStateMeta) & "|" & _
                TestMetaFingerprint(result.CrackWidthMeta) & "|" & TestMetaFingerprint(result.LongitudinalCrackMeta) & "|" & _
                TestMetaFingerprint(result.StabilityMeta) & vbLf
        Next i
        guard.Restore
    End If
    TestMeasurePerformanceBatch = "seconds=" & TestNumberText(total) & Chr$(30) & "preparation=" & TestNumberText(preparation) & Chr$(30) & _
        "core=" & TestNumberText(core) & Chr$(30) & "clear=" & TestNumberText(clear) & Chr$(30) & "packing=" & TestNumberText(packing) & Chr$(30) & _
        "summary=" & TestNumberText(summary) & Chr$(30) & "snapshot=" & TestNumberText(snapshot) & Chr$(30) & "plot=" & TestNumberText(plot) & _
        Chr$(30) & tech & Chr$(30) & "fingerprint=" & fingerprint & TestPerformanceResultsFingerprint()
    If mode = "Full" Then TestMeasurePerformanceBatch = TestMeasurePerformanceBatch & vbLf & TestPerformancePlotFingerprint()
    Exit Function
Failed:
    Dim number As Long, description As String
    number = Err.Number: description = Err.Description
    If Not guard Is Nothing Then guard.Restore
    Err.Raise number, "TestMeasurePerformanceBatch", description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает фактические данные схемы вне таймера: File не может ускориться за счёт пустого рисунка.
Private Function TestPerformancePlotFingerprint() As String
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    If Not settings.GetRequiredBoolean("Plot.Enabled") Then TestPerformancePlotFingerprint = "PLOT_DISABLED": Exit Function
    Dim reader As CSectionPlotDataReader: Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook ThisWorkbook, settings
    Dim text As String, lines() As String, i As Long
    text = "PLOT|" & reader.LoadCase & "|" & reader.StateType & "|" & reader.LengthUnit & "|" & reader.ResultUnit & "|" & CStr(reader.Count)
    ReDim lines(0 To reader.Count): lines(0) = text
    For i = 1 To reader.Count
        lines(i) = reader.ElementID(i) & "|" & TestDoubleBits(reader.X(i)) & "|" & TestDoubleBits(reader.Y(i)) & "|" & _
            TestDoubleBits(reader.Area(i)) & "|" & TestDoubleBits(reader.ResultValue(i)) & "|" & reader.PhysicalState(i)
    Next i
    TestPerformancePlotFingerprint = Join(lines, vbLf)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' один и тот же LC должен дать точные данные в пакете, повторном
' Execute, другом порядке и отдельном расчете. Адрес исходной строки сохраняется.
' Это проверка изоляции, а не замер ускорения 30 отдельных запусков.
Public Function RunPerformanceBatchIsolationTests(Optional ByVal family As String = "CRACK_SP35") As String
    Dim names As Variant, name As Variant, saved As Object, table As Object, values As Variant
    Dim batch As CBatchSectionCalculator, section As CSectionModel, reference As Object
    Dim i As Long, col As Long, order As Variant, singles() As Variant, loads As Object
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    Set saved = CreateObject("Scripting.Dictionary")
    names = Array("rngSystemSettings", "rngCalculationProfiles", "rngLoadCombinations", "rngStabilityDurationLoads")
    For Each name In names
        saved.Add CStr(name), ThisWorkbook.Names.Item(CStr(name)).RefersToRange.Formula
    Next name
    TestConfigurePerformanceFixture family, 30, "Newton"
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    values = loads.Offset(1, 0).Resize(loads.Rows.Count - 1, loads.Columns.Count).Value2
    Set reference = CreateObject("Scripting.Dictionary")
    Set batch = BuildPerformanceTestBatch(section): batch.Execute
    For i = 1 To batch.Count
        reference.Add batch.CombinationID(i), TestCombinationFingerprint(batch, section, i)
    Next i
    TestCheck "isolation.count", reference.Count = 30
    Dim retained As CCombinationResult, retainedText As String
    Set retained = batch.ResultAt(1): retainedText = TestResultFingerprint(retained, section)
    batch.Execute
    For i = 1 To batch.Count
        TestCheck "isolation.repeat." & batch.CombinationID(i), TestCombinationFingerprint(batch, section, i) = reference(batch.CombinationID(i))
    Next i
    TestCheck "isolation.retained", TestResultFingerprint(retained, section) = retainedText
#If PERFORMANCE_CURRENT Then
    TestCheck "isolation.released", batch.TestRunPreparationReleased
#End If
    ' Ошибка ввода содержит фактический адрес, который меняется при перестановке.
    ' Поэтому MIXED проверяется отдельно с сохранением исходного места каждой строки.
    If family <> "MIXED" Then
        For Each order In Array("Reverse", "Shuffle")
            TestConfigurePerformanceFixture family, 30, "Newton", "Saved", CStr(order)
            Set batch = BuildPerformanceTestBatch(section): batch.Execute
            For i = 1 To batch.Count
                TestCheck "isolation." & CStr(order) & "." & batch.CombinationID(i), _
                    TestCombinationFingerprint(batch, section, i) = reference(batch.CombinationID(i))
            Next i
        Next order
    End If
    For i = 1 To 30
        ReDim singles(1 To UBound(values, 1), 1 To UBound(values, 2))
        For col = 1 To UBound(values, 2): singles(i, col) = values(i, col): Next col
        loads.Offset(1, 0).Resize(UBound(singles, 1), UBound(singles, 2)).Value2 = singles
        Set batch = BuildPerformanceTestBatch(section): batch.Execute
        TestCheck "isolation.single.count." & CStr(i), batch.Count = 1
        TestCheck "isolation.single." & CStr(i), TestCombinationFingerprint(batch, section, 1) = reference(batch.CombinationID(1))
    Next i
    GoTo Restore
Failed:
    TestCheck "isolation.runtime " & CStr(Err.Number) & ": " & Err.Description, False
Restore:
    On Error Resume Next
    If Not saved Is Nothing Then
        For Each name In saved.Keys
            ThisWorkbook.Names.Item(CStr(name)).RefersToRange.Formula = saved(name)
        Next name
    End If
    On Error GoTo 0
    RunPerformanceBatchIsolationTests = mReport & "TOTAL PERFORMANCE ISOLATION " & family & ": passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' обычное чтение и подготовка объектов без записи Results и нового начального
' приближения.
Private Function BuildPerformanceTestBatch(ByRef section As CSectionModel) As CBatchSectionCalculator
    Dim settings As CSystemSettingsReader, units As CUnitSystem, provider As CMaterialModelProvider
    Dim profiles As CCalculationProfileCatalog, batch As CBatchSectionCalculator, reader As CLoadCombinationReader
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    Set batch = New CBatchSectionCalculator: batch.Initialize section, provider
    Set batch.ProfileCatalog = profiles: batch.ApplySettings settings, units
    Set reader = New CLoadCombinationReader: reader.LoadFromWorkbook ThisWorkbook, batch, units
    LoadStabilityDurationLoadsFromWorkbook ThisWorkbook, batch, units
    Dim data As Variant, inputError As String
    data = ReadSP35Table721FromWorkbook(ThisWorkbook, inputError): batch.SetSP35Table721 data, inputError
    Audit03ApplyLoadReferenceForTests section, settings, units, batch
    Set BuildPerformanceTestBatch = batch
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' рабочие нагрузки и все вложенные результаты одного сочетания без замеров времени.
Private Function TestCombinationFingerprint(ByVal batch As CBatchSectionCalculator, ByVal section As CSectionModel, ByVal index As Long) As String
    TestCombinationFingerprint = TestDoubleBits(batch.N(index)) & TestDoubleBits(batch.Mx(index)) & TestDoubleBits(batch.My(index)) & _
        TestDoubleBits(batch.UserMx(index)) & TestDoubleBits(batch.UserMy(index)) & TestResultFingerprint(batch.ResultAt(index), section)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' точные значения ширины трещин, несущей способности, устойчивости
' и каждого сохранённого НДС. Времена и счётчики проверяются отдельными замерами.
Private Function TestResultFingerprint(ByVal result As CCombinationResult, ByVal section As CSectionModel) As String
    Dim text As String, state As CSectionStateResult, i As Long
    text = TestMetaFingerprint(result.OverallMeta) & TestMetaFingerprint(result.DirectStateMeta) & TestMetaFingerprint(result.CapacityMeta) & _
        TestMetaFingerprint(result.CrackFormationMeta) & TestMetaFingerprint(result.CrackCurrentStateMeta) & TestMetaFingerprint(result.CrackWidthMeta) & _
        TestMetaFingerprint(result.LongitudinalCrackMeta) & TestMetaFingerprint(result.StabilityMeta)
    text = text & TestScalarFieldsFingerprint(result, "Status NormalCrackStatus CrackSummaryStatus ExtensionUsed GoverningUtil ProfileId DiagramExtensionEnabled")
    text = text & TestScalarFieldsFingerprint(result.StrengthResult.DirectState, "Status StateAvailable StateConverged ExtensionUsed ExtensionEnabled StrainReserve")
    text = text & TestScalarFieldsFingerprint(result.StrengthResult.Capacity, "Status LimitState SolutionMethod PathResolved LambdaCapacity UtilCapacity NUltimate MxUltimate MyUltimate CapacityStateN CapacityStateMx CapacityStateMy MomentUltimate CapacityStateAvailable CapacityEpsilon0 CapacityKappaX CapacityKappaY PureAxialForMethod ReserveFactor SearchExecuted")
    text = text & TestScalarFieldsFingerprint(result.CrackResult.Formation, "CrackFormed ConfirmedCracked ConfirmedNotCracked FormationMethod LambdaCrc Ncrc FormationNcrc Mcrc Ared RbtSer StopReason HasLimitPoint FallbackPsi1 CentralTensionBranch Converged")
    text = text & TestScalarFieldsFingerprint(result.CrackResult.Width, "CrackWidth AllowableCrackWidth Utilization Phi1 Phi2 Phi3 PsiS SigmaS SigmaSCrc SteelEs AsTension Abt DsEquivalent CrackSpacingRaw CrackSpacing TensionRebarCount TensionRebarIds SectionDepthH CoverA TensionDepth EffectiveZoneDepth CentralTensionBranch StopReason StandardCode SP35CriticalCandidate SP35ReinforcementRadius SP35PsiCm CriticalAnchorGroupID ReserveFactor")
    text = text & TestScalarFieldsFingerprint(result.CrackResult.Longitudinal, "Status MaxCompressionStress RbMc2 Utilization ReserveFactor")
    text = text & TestScalarFieldsFingerprint(result.StabilityResult, "Status SummaryReserve Code Branch ReferenceX ReferenceY DesignN DesignMx DesignMy SustainedN Moment1 Moment2 SustainedMoment1 SustainedMoment2 DesignMoment1 DesignMoment2 CriticalForce ReserveFactor Utilization Ncr1 Ncr2 Nultimate1 Nultimate2 Reserve1 Reserve2 Eta1 Eta2 PhiL1 PhiL2 PhiLTable1 PhiLTable2 PhiM1 PhiM2 PhiValue1 PhiValue2 PhiP1 PhiP2 Eccentricity1 Eccentricity2 AccidentalEcc1 AccidentalEcc2 EffectiveLength1 EffectiveLength2 Slenderness1 Slenderness2 Depth1 Depth2 Radius1 Radius2 CoreDistance1 CoreDistance2 Delta1 Delta2 Kb1 Kb2 Ks1 Ks2 StiffnessD1 StiffnessD2 PlaneBranch1 PlaneBranch2 PlaneApplicable1 PlaneApplicable2 PlanePassed1 PlanePassed2")
    For i = 1 To result.StateRepository.StateCount
        Set state = result.StateRepository.StateAt(i)
        text = text & TestScalarFieldsFingerprint(state, "StateType StateTypeText MaterialModelRole MaterialModelRoleText Epsilon0 KappaX KappaY ExtensionUsed WithinPhysicalRange Converged StopReason Status InternalStatus ResultCode ResultComment TargetN TargetMx TargetMy Nint Mxint Myint ResidualN ResidualMx ResidualMy RelativeResidualN RelativeResidualMx RelativeResidualMy LastResidualNorm IterationCount MinConcreteStrain MaxConcreteStrain MinSteelStrain MaxSteelStrain MinConcreteStress MaxConcreteStress MinSteelStress MaxSteelStress HasConcreteCompression HasConcreteTension HasSteelCompression HasSteelTension")
        text = text & state.MaterialSpec.SpecKey & TestArrayFingerprint(state.PrepareElementSnapshot(section))
    Next i
    Dim width As CCrackWidthResult, data As CSP35CrackData, j As Long
    Set width = result.CrackResult.Width
    If Not width Is Nothing Then
        text = text & TestRegionFingerprint(width.InteractionRegion)
        Set data = width.SP35Data
        If Not data Is Nothing Then
            text = text & TestScalarFieldsFingerprint(data, "CandidateCount GroupCount NormalX NormalY NeutralProjection GroupGapTolerance RowTolerance NeighborRatioLimit RadiusDiameterMode InteractionRadiusMode CentralTension SectionRevision")
            For j = 1 To data.GroupCount
                text = text & TestDoubleBits(data.GroupX(j)) & TestDoubleBits(data.GroupY(j)) & TestDoubleBits(data.GroupArea(j)) & _
                    TestDoubleBits(data.GroupBeta(j)) & CStr(data.GroupRow(j)) & TestArrayFingerprint(data.GroupMembers(j))
            Next j
            For j = 1 To data.CandidateCount
                text = text & CStr(data.CandidateAvailable(j))
                If Not data.CandidateAvailable(j) Then GoTo NextIsolationCandidate
                text = text & TestRegionFingerprint(data.CandidateRegion(j)) & TestArrayFingerprint(data.CandidateRows(j)) & _
                    CStr(data.CandidateAnchorGroup(j)) & CStr(data.CandidateStressBar(j)) & TestDoubleBits(data.CandidateSigmaS(j)) & _
                    TestDoubleBits(data.CandidateBetaDiameterSum(j)) & TestDoubleBits(data.CandidateSideDiameter(j)) & TestDoubleBits(data.CandidateRowDiameter(j)) & _
                    TestDoubleBits(data.CandidateLeftDistance(j)) & TestDoubleBits(data.CandidateRightDistance(j)) & TestDoubleBits(data.CandidateNormalRadius(j)) & _
                    CStr(data.CandidateReferenceRow(j)) & CStr(data.CandidateRowCount(j)) & TestDoubleBits(data.CandidateCrackWidth(j))
NextIsolationCandidate:
            Next j
        End If
    End If
    TestResultFingerprint = text
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' читает только явно перечисленные числовые и текстовые свойства без округления
' Double.
Private Function TestScalarFieldsFingerprint(ByVal value As Object, ByVal fields As String) As String
    If value Is Nothing Then TestScalarFieldsFingerprint = "<Nothing>": Exit Function
    Dim name As Variant, item As Variant
    For Each name In Split(fields, " ")
        item = CallByName(value, CStr(name), VbGet)
        TestScalarFieldsFingerprint = TestScalarFieldsFingerprint & CStr(name) & "=" & TestCellFingerprint(item) & "|"
    Next name
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет размерность и точные значения массива одного состояния.
Private Function TestArrayFingerprint(ByVal data As Variant) As String
    If Not IsArray(data) Then TestArrayFingerprint = TestCellFingerprint(data): Exit Function
    Dim dimensions As Long, row As Long, col As Long, item As Variant
    On Error Resume Next
    dimensions = UBound(data, 2)
    On Error GoTo 0
    If dimensions = 0 Then
        TestArrayFingerprint = "1D|" & CStr(LBound(data)) & "|" & CStr(UBound(data))
        For Each item In data: TestArrayFingerprint = TestArrayFingerprint & TestCellFingerprint(item) & "|": Next item
    Else
        TestArrayFingerprint = "2D|" & CStr(UBound(data, 1)) & "|" & CStr(UBound(data, 2))
        For row = LBound(data, 1) To UBound(data, 1)
            For col = LBound(data, 2) To UBound(data, 2)
                TestArrayFingerprint = TestArrayFingerprint & TestCellFingerprint(data(row, col)) & "|"
            Next col
        Next row
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' статус, код, применимость и комментарий без подмены внешним текстом.
Private Function TestMetaFingerprint(ByVal meta As CResultMeta) As String
    TestMetaFingerprint = CStr(meta.InternalStatus) & ":" & CStr(meta.ResultCode) & ":" & CStr(meta.Applies) & ":" & _
        CStr(meta.Calculated) & ":" & meta.ResultComment
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' числа как биты Double, остальные значения - точный текст/тип.
' Исключены только RunID и summary solves/time. RunID проверяется внутри снимка.
Public Function TestPerformanceResultsFingerprint() As String
    Dim names As Variant, name As Variant, anchor As Object, data As Variant
    Dim rows As Long, columns As Long, row As Long, col As Long, result As String, value As Variant
    Dim lines() As String, line As String
    Dim ndm As CNDMResultsWriter, runID As String
    Set ndm = New CNDMResultsWriter
    names = Array("rngBatchSummary", "rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor", _
        "rngNDMSectionGeometry", "rngNDMSectionContours", "rngNDMElementResults", "rngNDMSectionProperties", _
        "rngNDMMaterialDiagrams", "rngNDMSectionAnnotations")
    For Each name In names
        Set anchor = ThisWorkbook.Names.Item(CStr(name)).RefersToRange.Cells(1, 1)
        If name = "rngBatchSummary" Then
            rows = 42: columns = 26
        ElseIf name = "rngStrengthSummaryAnchor" Then
            rows = 30: columns = 49
        ElseIf name = "rngCrackSummaryAnchor" Then
            Dim cw As CCrackSummaryWriter: Set cw = New CCrackSummaryWriter
            rows = 30: columns = cw.RequiredColumns
        ElseIf name = "rngStabilitySummaryAnchor" Then
            Dim sw As CStabilitySummaryWriter: Set sw = New CStabilitySummaryWriter
            rows = 30: columns = sw.RequiredColumns
        Else
            rows = TestOldExtent(anchor, False, 0): columns = ndm.OutputColumnCount(CStr(name))
            If rows < 1 Then rows = 1
        End If
        data = anchor.Resize(rows, columns).Value2
        If CStr(name) = "rngNDMElementResults" Or CStr(name) = "rngNDMSectionProperties" Or CStr(name) = "rngNDMSectionAnnotations" Then
            data = ReadAnchoredResultTable(ThisWorkbook, CStr(name)): rows = UBound(data, 1): columns = UBound(data, 2)
        End If
        ReDim lines(0 To rows)
        lines(0) = CStr(name) & "|" & CStr(rows) & "|" & CStr(columns)
        For row = 1 To rows
            line = vbNullString
            For col = 1 To columns
                value = data(row, col)
                If Left$(CStr(name), 6) = "rngNDM" And col = 1 And row > 1 Then
                    If Len(runID) = 0 Then runID = CStr(value)
                    If CStr(value) <> runID Then Err.Raise 5, "TestPerformanceResultsFingerprint", "Inconsistent RunID."
                    value = "<RunID>"
                End If
                If name = "rngBatchSummary" And col = 1 And (row = 3 Or row = 4) Then value = "<Technical>"
                line = line & TestCellFingerprint(value) & "|"
            Next col
            lines(row) = line
        Next row
        If CStr(name) = "rngNDMSectionProperties" Or CStr(name) = "rngNDMSectionAnnotations" Then
            If rows > 2 Then TestSortFingerprintLines lines, 2, rows ' Общие и LC-строки физически разделены; сравниваются все значения независимо от порядка блоков.
        End If
        result = result & Join(lines, vbLf) & vbLf
    Next name
    TestPerformanceResultsFingerprint = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Упорядочивает только сериализованные строки свойств/аннотаций; ни одно поле не исключает.
Private Sub TestSortFingerprintLines(ByRef lines() As String, ByVal first As Long, ByVal last As Long)
    Dim left As Long, right As Long, pivot As String, temporary As String
    left = first: right = last: pivot = lines((first + last) \ 2)
    Do While left <= right
        Do While StrComp(lines(left), pivot, vbBinaryCompare) < 0: left = left + 1: Loop
        Do While StrComp(lines(right), pivot, vbBinaryCompare) > 0: right = right - 1: Loop
        If left <= right Then
            temporary = lines(left): lines(left) = lines(right): lines(right) = temporary
            left = left + 1: right = right - 1
        End If
    Loop
    If first < right Then TestSortFingerprintLines lines, first, right
    If left < last Then TestSortFingerprintLines lines, left, last
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' различает Empty/error/типы и сохраняет биты числовой ячейки.
Private Function TestCellFingerprint(ByVal value As Variant) As String
    If IsError(value) Then
        TestCellFingerprint = "E:" & CStr(value)
    ElseIf IsEmpty(value) Or (VarType(value) = vbString And Len(CStr(value)) = 0) Then
        TestCellFingerprint = "Empty"
    ElseIf VarType(value) = vbBoolean Then
        TestCellFingerprint = "T:8:" & CStr(Len(CStr(value))) & ":" & CStr(value) ' Excel преобразует текст True/False в Boolean; сохранённый смысл флага одинаков.
    ElseIf VarType(value) = vbDouble Then
        TestCellFingerprint = "D:" & TestDoubleBits(CDbl(value))
    Else
        TestCellFingerprint = "T:" & CStr(VarType(value)) & ":" & CStr(Len(CStr(value))) & ":" & CStr(value)
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' повторяет обычное чтение сохранённых данных экспорта без запуска
' AutoCAD. Время включает геометрию и выбор сочетания; проверки выполняются вне замера.
Public Function TestMeasurePerformanceExport(ByVal count As Long) As String
    Dim t As Double, i As Long, elapsed As Double, values() As String
    ReDim values(1 To count)
    t = TestPerformanceNow()
    For i = 1 To count
        values(i) = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    Next i
    elapsed = TestSecondsSince(t)
    For i = 2 To count
        If values(i) <> values(1) Then Err.Raise 5, "TestMeasurePerformanceExport", "Read changed within measurement."
    Next i
    TestMeasurePerformanceExport = "seconds=" & TestNumberText(elapsed) & Chr$(30) & "fingerprint=" & values(1) & vbLf & TestPerformanceResultsFingerprint()
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' следующая операция видит новые данные; ошибка не удерживает кэш.
Public Function RunPerformanceExportCacheTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    Dim table As Object, data As Variant, row As Long, original As String, changed As String
    Dim oldValue As Variant, oldHeader As Variant, errorNumber As Long, restored As Boolean
    On Error GoTo Failed
    Set table = ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange
    Set table = table.Resize(TestOldExtent(table.Cells(1, 1), False, 0), 9)
    data = table.Value2
    original = Audit02ReadExportSnapshotForTests(ThisWorkbook)
#If PERFORMANCE_CURRENT Then
    TestCheck "export.released", InStr(PerformanceExportReadDiagnosticsForTests(), "released=True") > 0
#End If
    Dim identity As Variant
    identity = Split(original, "|")
    For row = 2 To UBound(data, 1)
        If CStr(data(row, 2)) = identity(0) And CStr(data(row, 4)) = identity(2) Then Exit For
    Next row
    If row > UBound(data, 1) Then Err.Raise 5, "RunPerformanceExportCacheTests", "Missing selected state row."
    oldValue = table.Cells(row, 8).Formula
    table.Cells(row, 8).Value2 = CDbl(data(row, 8)) + 0.123456789
    changed = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    TestCheck "export.fresh.operation", changed <> original
    table.Cells(row, 8).Formula = oldValue: restored = True
    TestCheck "export.restored", Audit02ReadExportSnapshotForTests(ThisWorkbook) = original
    oldHeader = table.Cells(1, 8).Formula
    table.Cells(1, 8).Value2 = "broken"
    On Error Resume Next
    changed = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    errorNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "export.corrupt.rejected", errorNumber <> 0
#If PERFORMANCE_CURRENT Then
    TestCheck "export.failure.released", InStr(PerformanceExportReadDiagnosticsForTests(), "released=True") > 0
#End If
    table.Cells(1, 8).Formula = oldHeader
    TestCheck "export.after.failure", Audit02ReadExportSnapshotForTests(ThisWorkbook) = original
    GoTo Restore
Failed:
    TestCheck "export.runtime " & CStr(Err.Number) & ": " & Err.Description, False
Restore:
    On Error Resume Next
    If Not restored And row > 1 Then table.Cells(row, 8).Formula = oldValue
    If Not IsEmpty(oldHeader) Then table.Cells(1, 8).Formula = oldHeader
    On Error GoTo 0
    RunPerformanceExportCacheTests = mReport & "TOTAL PERFORMANCE EXPORT CACHE: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отделяет B04 от дорогого solve, а P05 от записи чисел.
' Фактическая стоимость 30 вызовов оценивается по серии 10000, не по нулевому Timer.
Public Function TestMeasurePerformanceOverhead() As String
    Dim catalog As CCalculationProfileCatalog, profile As CCalculationProfile, meta As CResultMeta, copy As CResultMeta
    Dim i As Long, t As Double, validation As Double, cloning As Double, formatting As Double, errorText As String
    Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
    Set profile = catalog.ProfileById("PR1")
    Set meta = New CResultMeta: meta.SetResult rsSuccess, rcCheckPassed, rkDirectState, "Расчет выполнен."
    t = TestPerformanceNow()
    For i = 1 To 10000
        If Not profile.ValidateForCalculation(errorText) Then Err.Raise 5, "TestMeasurePerformanceOverhead", errorText
    Next i
    validation = TestSecondsSince(t)
    t = TestPerformanceNow()
    For i = 1 To 10000: Set copy = meta.Clone: Next i
    cloning = TestSecondsSince(t)
    Dim names As Variant, name As Variant, ndm As CNDMResultsWriter, anchor As Object, rows As Long, columns As Long, title As String
    Set ndm = New CNDMResultsWriter
    names = Array("rngNDMSectionGeometry", "rngNDMSectionContours", "rngNDMElementResults", "rngNDMSectionProperties", _
        "rngNDMMaterialDiagrams", "rngNDMSectionAnnotations")
    For Each name In names
        Set anchor = ThisWorkbook.Names.Item(CStr(name)).RefersToRange.Cells(1, 1)
        rows = TestOldExtent(anchor, False, 0): If rows < 1 Then rows = 1
        columns = ndm.OutputColumnCount(CStr(name)): title = CStr(anchor.Offset(-1, 0).Value2)
        t = TestPerformanceNow(): ndm.FormatResultsBlock anchor, rows, columns, title
        formatting = formatting + TestSecondsSince(t)
    Next name
    TestMeasurePerformanceOverhead = "validation10000=" & TestNumberText(validation) & Chr$(30) & "clone10000=" & TestNumberText(cloning) & _
        Chr$(30) & "snapshotFormat=" & TestNumberText(formatting)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' готовит одинаковый сохраненный НДС до замера настоящего CAD.
' Настройки меняются только в частной книге runner-а; расчет не входит во время
' импорта/экспорта. Повернутый вариант проверяет повторяющиеся WCS-оболочки.
Public Sub TestPrepareNativeCADPerformance(ByVal rotated As Boolean)
    TestConfigurePerformanceFixture "IMPORTED_DIRECT", 1, "Newton"
    TestPerformanceSetting "Plot.Enabled", "No"
    TestPerformanceSetting "Plot.AutoUpdateAfterCalculation", "No"
    TestPerformanceSetting "General.ExecutionReportEnabled", "No"
    TestPerformanceSetting "AutoCAD.Export.CombinationID", "PERF_01"
    TestPerformanceSetting "AutoCAD.Import.MinArea", "0"
    Dim table As Object, row As Long
    Set table = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    For row = 1 To table.Rows.Count
        If CStr(table.Cells(row, 2).Value2) = "Visualization.State" Then table.Cells(row, 3).Value2 = "StrengthState"
    Next row
    If rotated Then
        Dim original As CSectionModel, section As CSectionModel, i As Long, c As Double, s As Double
        Dim query As CSectionGeometryQuery, region As CConcreteRegion, settings As CSystemSettingsReader, units As CUnitSystem
        Dim writer As CNDMResultsWriter, x As Double, y As Double
        Set original = ReadSectionGeometryFromResults(ThisWorkbook)
        Set section = New CSectionModel: section.SourceType = "AutoCADImport"
        c = Cos(0.37): s = Sin(0.37)
        For i = 1 To original.ConcreteCount
            x = original.ConcreteX(i): y = original.ConcreteY(i)
            section.AddConcreteElement c * x - s * y, s * x + c * y, original.ConcreteArea(i), 1, , , "Rectangle", 10#, 10#, 0.37
        Next i
        For i = 1 To original.RebarCount
            x = original.RebarX(i): y = original.RebarY(i)
            section.AddRebarElement c * x - s * y, s * x + c * y, original.RebarDiameter(i), original.RebarArea(i), "Periodic"
        Next i
        Set query = New CSectionGeometryQuery: Set region = query.CircleRegion(0#, 0#, 200#)
        Set region = query.SubtractRegion(region, query.CircleRegion(c * -45# - s * 60#, s * -45# + c * 60#, 25#))
        Set region = query.SubtractRegion(region, query.CircleRegion(c * 45# - s * 60#, s * 45# + c * 60#, 25#))
        section.Contours.AddMaterialRegion region, "PERFORMANCE_CAD", "Повернутый достоверный контур теста."
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set writer = New CNDMResultsWriter
        writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section), units
    End If
    Dim result As String
    result = TestMeasurePerformanceBatch("Staged")
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' измеряет настоящий экспорт и импорт одного нового документа.
' Сигнатура содержит все расчетные геометрические поля, без разных CAD handles;
' подписи, слои и цвета проверяются вне таймеров. Пользовательский DWG запрещен.
Public Function TestMeasureNativeCADPerformance(ByVal doc As Object, ByVal noiseCount As Long) As String
    If InStr(1, CStr(doc.Application.FullName), "\AutoCAD 2023\acad.exe", vbTextCompare) = 0 Then Err.Raise 5, , "Требуется настоящий Autodesk AutoCAD."
    If doc.ModelSpace.Count <> 0 Then Err.Raise 5, , "Замер разрешен только в пустом собственном документе."
    Dim t As Double, exportSeconds As Double, importSeconds As Double, count As Long, solves As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, importer As CAutoCADSectionModelImporter
    Dim section As CSectionModel, entity As Object, i As Long, signature As String
    Dim userLine As Object, userSet As Object, userPoint(0 To 2) As Double, userEnd(0 To 2) As Double
    Dim userLayer As Object, userHandle As String, selectionCount As Long
    Set userLayer = doc.Layers.Add("PERFORMANCE_EXISTING")
    userPoint(0) = 1000#: userEnd(0) = 1010#: userEnd(1) = 10#
    Set userLine = doc.ModelSpace.AddLine(userPoint, userEnd)
    userLine.Layer = userLayer.Name: userLine.Color = 173: userHandle = CStr(userLine.Handle)
    Set userSet = doc.SelectionSets.Add("PERFORMANCE_USER_SELECTION")
    Dim userItems(0 To 0) As Object: Set userItems(0) = userLine
    userSet.AddItems userItems: selectionCount = doc.SelectionSets.Count
    solves = SectionEquilibriumSolveCount()
    t = TestPerformanceNow(): count = Audit03ExportAutoCADDocumentForTests(doc): exportSeconds = TestSecondsSince(t)
    If CStr(userLine.Handle) <> userHandle Or CStr(userLine.Layer) <> "PERFORMANCE_EXISTING" Or userLine.Color <> 173 Then Err.Raise 5, , "Экспорт изменил прежний объект чертежа."
    Dim savedPoint As Variant: savedPoint = userLine.EndPoint
    If CDbl(savedPoint(0)) <> 1010# Or CDbl(savedPoint(1)) <> 10# Then Err.Raise 5, , "Экспорт изменил координаты прежнего объекта."
    If doc.SelectionSets.Count <> selectionCount Or userSet.Count <> 1 Then Err.Raise 5, , "Экспорт изменил прежние наборы выбора или оставил временный набор."
    For Each entity In doc.ModelSpace
        signature = signature & CStr(entity.ObjectName) & "|" & CStr(entity.Layer) & "|" & CStr(entity.Color)
        If entity.ObjectName = "AcDbText" Then signature = signature & "|" & CStr(entity.TextString)
        signature = signature & vbLf
    Next entity
    If noiseCount > 0 Then
        Dim layer As Object, p(0 To 2) As Double, q(0 To 2) As Double
        Set layer = doc.Layers.Add("PERFORMANCE_UNRELATED")
        For i = 1 To noiseCount
            p(0) = 1000# + i: q(0) = p(0): q(1) = 10#
            Set entity = doc.ModelSpace.AddLine(p, q): entity.Layer = layer.Name
        Next i
    End If
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set importer = New CAutoCADSectionModelImporter
    Dim before As Long: before = doc.ModelSpace.Count
    t = TestPerformanceNow()
    Set section = importer.ImportConfiguredModelSpace(doc.ModelSpace, settings, units, doc.Layers)
    importSeconds = TestSecondsSince(t)
    If doc.ModelSpace.Count <> before Then Err.Raise 5, , "Импорт изменил исходные CAD-объекты."
    If SectionEquilibriumSolveCount() <> solves Then Err.Raise 5, , "Импорт/экспорт неожиданно запустил решатель."
    For i = 1 To section.ConcreteCount
        signature = signature & "C|" & section.ConcreteID(i) & "|" & section.ConcreteShapeType(i)
        signature = signature & "|" & TestDoubleBits(section.ConcreteX(i)) & "|" & TestDoubleBits(section.ConcreteY(i)) & "|" & TestDoubleBits(section.ConcreteArea(i))
        signature = signature & "|" & TestDoubleBits(section.ConcreteWidth(i)) & "|" & TestDoubleBits(section.ConcreteHeight(i)) & "|" & TestDoubleBits(section.ConcreteRotation(i))
        signature = signature & "|" & TestDoubleBits(section.ConcreteLocalIx(i)) & "|" & TestDoubleBits(section.ConcreteLocalIy(i)) & "|" & TestDoubleBits(section.ConcreteLocalIxy(i)) & vbLf
    Next i
    For i = 1 To section.RebarCount
        signature = signature & "R|" & section.RebarID(i) & "|" & TestDoubleBits(section.RebarX(i)) & "|" & TestDoubleBits(section.RebarY(i))
        signature = signature & "|" & TestDoubleBits(section.RebarArea(i)) & "|" & TestDoubleBits(section.RebarDiameter(i)) & vbLf
    Next i
    For i = 1 To section.Contours.Count
        signature = signature & "CONTOUR|" & section.Contours.SegmentType(i) & "|" & section.Contours.LoopRole(i)
        signature = signature & "|" & TestDoubleBits(section.Contours.StartX(i)) & "|" & TestDoubleBits(section.Contours.StartY(i))
        signature = signature & "|" & TestDoubleBits(section.Contours.EndX(i)) & "|" & TestDoubleBits(section.Contours.EndY(i))
        signature = signature & "|" & TestDoubleBits(section.Contours.CenterX(i)) & "|" & TestDoubleBits(section.Contours.CenterY(i))
        signature = signature & "|" & TestDoubleBits(section.Contours.Radius(i)) & "|" & TestDoubleBits(section.Contours.SweepAngle(i)) & vbLf
    Next i
    TestMeasureNativeCADPerformance = "export=" & TestNumberText(exportSeconds) & Chr$(30) & "import=" & TestNumberText(importSeconds) & _
        Chr$(30) & "entities=" & CStr(count) & Chr$(30) & "concrete=" & CStr(section.ConcreteCount) & Chr$(30) & _
        "rebar=" & CStr(section.RebarCount) & Chr$(30) & "noise=" & CStr(noiseCount) & Chr$(30) & "signature=" & signature
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' повёрнутое прямоугольное сечение с разными главными инерциями.
' Неравенство сторон исключает неопределённость направления осей круглого сечения.
Public Sub TestPrepareNativeCADRoundoffFixture()
    TestConfigurePerformanceFixture "IMPORTED_DIRECT", 1, "Newton"
    TestPerformanceSetting "Plot.Enabled", "No": TestPerformanceSetting "Plot.AutoUpdateAfterCalculation", "No"
    TestPerformanceSetting "General.ExecutionReportEnabled", "No"
    TestPerformanceSetting "AutoCAD.Export.CombinationID", "PERF_01"
    TestPerformanceSetting "AutoCAD.Import.MinArea", "0"
    Dim section As CSectionModel, i As Long, j As Long, x As Double, y As Double, c As Double, s As Double
    Dim points As Variant, point As Variant, path As CGeometryPath, query As CSectionGeometryQuery
    Dim settings As CSystemSettingsReader, units As CUnitSystem, writer As CNDMResultsWriter, table As Object
    Set section = New CSectionModel: section.SourceType = "AutoCADImport"
    c = Cos(0.37): s = Sin(0.37)
    For i = 1 To 8
        For j = 1 To 6
            x = -300# + (i - 0.5) * 75#: y = -150# + (j - 0.5) * 50#
            section.AddConcreteElement c * x - s * y, s * x + c * y, 3750#, 1, , , "Rectangle", 75#, 50#, 0.37
        Next j
    Next i
    points = Array(Array(-260#, -110#), Array(0#, -110#), Array(260#, -110#), Array(260#, 0#), _
        Array(260#, 110#), Array(0#, 110#), Array(-260#, 110#), Array(-260#, 0#))
    For Each point In points
        x = CDbl(point(0)): y = CDbl(point(1))
        section.AddRebarElement c * x - s * y, s * x + c * y, 20#, 0#, "Periodic"
    Next point
    Set path = New CGeometryPath: path.Initialize "Rectangle"
    path.AddLine "Bottom", -300#, -150#, 300#, -150#: path.AddLine "Right", 300#, -150#, 300#, 150#
    path.AddLine "Top", 300#, 150#, -300#, 150#: path.AddLine "Left", -300#, 150#, -300#, -150#: path.Seal True
    Set query = New CSectionGeometryQuery
    section.Contours.AddMaterialRegion query.RegionFromPath(path.Transformed(0#, 0#, 0.37)), "ROUNDING_TEST", "Контур контрольного прямоугольника."
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section), units
    Set table = ThisWorkbook.Names("rngCalculationProfiles").RefersToRange
    For i = 1 To table.Rows.Count
        If CStr(table.Cells(i, 2).Value2) = "Visualization.State" Then table.Cells(i, 3).Value2 = "StrengthState"
    Next i
    Dim measured As String: measured = TestMeasurePerformanceBatch("Staged")
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' одна и та же рабочая математика для двух реально импортированных чертежей.
' Проверяет прочность, трещины и устойчивость; нагрузки заданы в Н и Н*мм.
Public Function TestMeasureCADRoundoffCalculation(ByVal document As Object, ByVal extension As String) As String
    Dim areas As Variant, name As Variant, snapshots As Object, area As Object
    Dim number As Long, owner As String, reason As String, report As String
    Set snapshots = CreateObject("Scripting.Dictionary")
    areas = Array("rngSystemSettings", "rngCalculationProfiles", "rngLoadCombinations", "rngStabilityDurationLoads")
    For Each name In areas
        snapshots.Add CStr(name), ThisWorkbook.Names(CStr(name)).RefersToRange.Formula
    Next name
    On Error GoTo Failed
    TestConfigurePerformanceFixture "MIXED", 6, "Newton", "Saved", "Normal", extension
    Dim settings As CSystemSettingsReader, units As CUnitSystem, importer As CAutoCADSectionModelImporter
    Dim section As CSectionModel, provider As CMaterialModelProvider, profiles As CCalculationProfileCatalog
    Dim batch As CBatchSectionCalculator, loads As Variant, load As Variant, i As Long, j As Long
    Dim result As CCombinationResult, state As CSectionStateResult, data As Variant, inputError As String, converged As Long
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set importer = New CAutoCADSectionModelImporter
    Set section = importer.ImportConfiguredModelSpace(document.ModelSpace, settings, units, document.Layers)
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    Set batch = New CBatchSectionCalculator: batch.Initialize section, provider
    Set batch.ProfileCatalog = profiles: batch.ApplySettings settings, units
    loads = Array(Array(-500000#, 0#, 0#), Array(0#, 50000000#, 0#), Array(-500000#, 50000000#, -20000000#), _
        Array(200000#, 20000000#, 10000000#), Array(-5000000#, 500000000#, 200000000#), Array(-200000#, -30000000#, 40000000#))
    For Each load In loads
        i = i + 1
        batch.AddCombination "ROUND_" & CStr(i), CDbl(load(0)), CDbl(load(1)), CDbl(load(2)), "PR1", "Контроль погрешности AutoCAD."
        batch.AddStabilityDurationLoad "ROUND_" & CStr(i), 0#, 0#, 0#
    Next load
    data = ReadSP35Table721FromWorkbook(ThisWorkbook, inputError): batch.SetSP35Table721 data, inputError
    Audit03ApplyLoadReferenceForTests section, settings, units, batch
    batch.Execute
    For i = 1 To batch.Count
        Set result = batch.ResultAt(i)
        report = report & TestCADRoundoffFields(CStr(i) & ".Result", result, "Status NormalCrackStatus CrackSummaryStatus GoverningUtil")
        report = report & TestCADRoundoffFields(CStr(i) & ".Meta", result.OverallMeta, "InternalStatus ResultCode Applies Calculated")
        report = report & TestCADRoundoffFields(CStr(i) & ".Direct", result.StrengthResult.DirectState, "Status StateAvailable StateConverged StrainReserve")
        report = report & TestCADRoundoffFields(CStr(i) & ".Capacity", result.StrengthResult.Capacity, "Status LimitState LambdaCapacity UtilCapacity NUltimate MxUltimate MyUltimate CapacityStateAvailable CapacityEpsilon0 CapacityKappaX CapacityKappaY")
        report = report & TestCADRoundoffFields(CStr(i) & ".Formation", result.CrackResult.Formation, "CrackFormed ConfirmedCracked ConfirmedNotCracked LambdaCrc Ncrc Mcrc Converged StopReason")
        report = report & TestCADRoundoffFields(CStr(i) & ".Width", result.CrackResult.Width, "CrackWidth Utilization PsiS SigmaS SigmaSCrc CrackSpacing StopReason StandardCode")
        report = report & TestCADRoundoffFields(CStr(i) & ".Longitudinal", result.CrackResult.Longitudinal, "Status MaxCompressionStress Utilization")
        report = report & TestCADRoundoffFields(CStr(i) & ".Stability", result.StabilityResult, "Status SummaryReserve CriticalForce Ncr1 Ncr2 Reserve1 Reserve2 Eta1 Eta2")
        report = report & CStr(i) & ".Repository|Count|T|" & CStr(result.StateRepository.StateCount) & vbLf
        If result.StrengthResult.DirectState.StateConverged Then converged = converged + 1
        For j = 1 To result.StateRepository.StateCount
            Set state = result.StateRepository.StateAt(j)
            report = report & TestCADRoundoffFields(CStr(i) & ".State." & CStr(j), state, "StateType Converged Status InternalStatus ResultCode Epsilon0 KappaX KappaY TargetN TargetMx TargetMy Nint Mxint Myint MinConcreteStrain MaxConcreteStrain MinSteelStrain MaxSteelStrain MinConcreteStress MaxConcreteStress MinSteelStress MaxSteelStress")
        Next j
    Next i
    If batch.Count <> 6 Or converged < 3 Then Err.Raise 5, "CADRoundoff", "Недостаточно пригодных контрольных результатов."
    report = report & "Control|Cases|T|6" & vbLf
    GoTo Restore
Failed:
    number = Err.Number: owner = Err.Source: reason = Err.Description
Restore:
    For Each name In areas
        ThisWorkbook.Names(CStr(name)).RefersToRange.Formula = snapshots.Item(CStr(name))
    Next name
    If number <> 0 Then Err.Raise number, owner, reason
    TestMeasureCADRoundoffCalculation = report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет исходные биты чисел и точные статусы без округления результатов.
Public Function TestCADRoundoffFields(ByVal prefix As String, ByVal target As Object, ByVal fields As String) As String
    If target Is Nothing Then TestCADRoundoffFields = prefix & "|Available|T|False" & vbLf: Exit Function
    Dim name As Variant, value As Variant, kind As String, text As String
    For Each name In Split(fields, " ")
        value = CallByName(target, CStr(name), VbGet)
        If VarType(value) = vbDouble Then
            kind = "D": text = TestDoubleBits(CDbl(value))
        Else
            kind = "T": text = Replace(Replace(CStr(value), "|", "%7C"), vbCrLf, "%0A")
        End If
        TestCADRoundoffFields = TestCADRoundoffFields & prefix & "|" & CStr(name) & "|" & kind & "|" & text & vbLf
    Next name
End Function
