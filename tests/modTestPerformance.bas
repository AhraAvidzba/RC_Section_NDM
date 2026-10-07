Attribute VB_Name = "modTestPerformance"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: ЭКВИВАЛЕНТНОСТЬ И ЗАМЕРЫ ЛОКАЛЬНЫХ ОПТИМИЗАЦИЙ
' ==========================================================================
' Стандартный модуль не добавляет production-классов. Проверяет реальные
' контейнеры и Excel Range, сохраняет точные значения без новых допусков.
' Замеры не включают assertions; тот же код импортируется в A и B книги.

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

' Проверяет рост/очистку массивов и прежнее правило первой пустой ячейки.
Public Function RunPerformanceStorageTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    TestModelStorage
    TestSnapshotExtent
    GoTo Finished
Failed:
    Check "storage.runtime " & CStr(Err.Number) & ": " & Err.Description, False
Finished:
    RunPerformanceStorageTests = mReport & "TOTAL PERFORMANCE STORAGE: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' Измеряет действующий Add модели. Числа результата проверяются внешним A/B;
' индекс/площадь/суммы считаются после таймера, не входят в preparationSeconds.
Public Function MeasurePerformanceModel(ByVal count As Long) As String
    Dim model As CSectionModel, i As Long, t As Double, elapsed As Double
    Set model = New CSectionModel
    t = PerformanceNow()
    For i = 1 To count
        model.AddConcreteElement CDbl(i), -CDbl(i), 100#, 1, "source" & CStr(i), vbNullString, "Rectangle", 10#, 10#
    Next i
    elapsed = SecondsSince(t)
    Dim area As Double, x As Double, y As Double
    For i = 1 To model.ConcreteCount
        area = area + model.ConcreteArea(i)
        x = x + model.ConcreteX(i): y = y + model.ConcreteY(i)
    Next i
    MeasurePerformanceModel = "seconds=" & NumberText(elapsed) & "; count=" & CStr(model.ConcreteCount) & _
        "; area=" & NumberText(area) & "; x=" & NumberText(x) & "; y=" & NumberText(y) & _
        "; last=" & model.ConcreteID(count)
End Function

' Сравнивает старое поячеечное правило и production-блочное чтение на одном
' Range. Все данные и размер одинаковы; test sheet удаляется даже при ошибке.
Public Function MeasurePerformanceExtent(ByVal optimized As Boolean, ByVal count As Long) As String
    On Error GoTo Failed
    Dim ws As Object, values() As Variant, i As Long, t As Double, elapsed As Double, found As Long
    Set ws = ThisWorkbook.Worksheets.Add
    ReDim values(1 To count, 1 To 1)
    For i = 1 To count: values(i, 1) = "row" & CStr(i): Next i
    ws.Range("A1").Resize(count, 1).Value2 = values
    t = PerformanceNow()
    If optimized Then
        found = CLng(Application.Run("'" & ThisWorkbook.Name & "'!modWorkbookCalculation.SnapshotAnchoredExtent", ws.Range("A1"), False))
    Else
        found = OldExtent(ws.Range("A1"), False, 0)
    End If
    elapsed = SecondsSince(t)
    MeasurePerformanceExtent = "seconds=" & NumberText(elapsed) & "; count=" & CStr(found)
    ws.Delete
    Exit Function
Failed:
    Dim errorNumber As Long, description As String
    errorNumber = Err.Number: description = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo 0
    Err.Raise errorNumber, "MeasurePerformanceExtent", description
End Function

' ДЛЯ ТЕСТОВ: поля по обе стороны нескольких границ резерва и Clear.
Private Sub TestModelStorage()
    Dim model As CSectionModel, i As Long, revision As Long, value As Double, errorNumber As Long
    Set model = New CSectionModel
    For i = 1 To 257
        model.AddConcreteElement CDbl(i) / 8#, -CDbl(i) / 16#, CDbl(i) + 1#, 2, "source" & CStr(i), "handle" & CStr(i)
        model.AddRebarElement CDbl(i) / 32#, CDbl(i) / 64#, 20#, CDbl(i) + 2#, "Periodic", 3, "bar" & CStr(i), "r" & CStr(i)
    Next i
    Check "model.count", model.ConcreteCount = 257 And model.RebarCount = 257
    For i = 1 To 257
        Check "model.concrete." & CStr(i), model.ConcreteID(i) = "C" & CStr(i) And _
            model.ConcreteX(i) = CDbl(i) / 8# And model.ConcreteY(i) = -CDbl(i) / 16# And _
            model.ConcreteArea(i) = CDbl(i) + 1# And model.ConcreteSourceHandle(i) = "handle" & CStr(i)
        Check "model.rebar." & CStr(i), model.RebarID(i) = "R" & CStr(i) And _
            model.RebarX(i) = CDbl(i) / 32# And model.RebarY(i) = CDbl(i) / 64# And _
            model.RebarArea(i) = CDbl(i) + 2# And model.RebarSourceName(i) = "bar" & CStr(i)
    Next i
    revision = model.Revision
    On Error Resume Next
    model.AddConcreteElement 0#, 0#, 0#: errorNumber = Err.Number: Err.Clear
    On Error GoTo 0
    Check "model.invalidConcrete", errorNumber <> 0 And model.ConcreteCount = 257 And model.Revision = revision
    On Error Resume Next
    model.AddRebarElement 0#, 0#, 0#, 0#, "Periodic": errorNumber = Err.Number: Err.Clear
    On Error GoTo 0
    Check "model.invalidRebar", errorNumber <> 0 And model.RebarCount = 257 And model.Revision = revision
    On Error Resume Next
    value = model.ConcreteX(258): errorNumber = Err.Number: Err.Clear
    On Error GoTo 0
    Check "model.tailHidden", errorNumber <> 0
    model.Clear
    Check "model.clear", model.ConcreteCount = 0 And model.RebarCount = 0 And model.Revision = revision + 1
    model.AddConcreteElement 1#, 2#, 3#: model.AddRebarElement 4#, 5#, 6#, 7#, "Smooth"
    Check "model.rebuild", model.ConcreteID(1) = "C1" And model.ConcreteX(1) = 1# And model.RebarID(1) = "R1" And model.RebarArea(1) = 7#
End Sub

' ДЛЯ ТЕСТОВ: Excel scalar/array, Empty/formula/space/error и край листа.
Private Sub TestSnapshotExtent()
    On Error GoTo Failed
    Dim ws As Object, anchor As Object, values() As Variant, i As Long
    Set ws = ThisWorkbook.Worksheets.Add
    ReDim values(1 To 2050, 1 To 1)
    For i = 1 To 2050: values(i, 1) = "r" & CStr(i): Next i
    ws.Range("B2").Resize(2050, 1).Value2 = values
    CheckExtent ws.Range("B2"), False, 0, "rows.blocks"
    ws.Range("B1027").Value2 = "   "
    CheckExtent ws.Range("B2"), False, 0, "rows.space"
    ws.Range("B1027").Formula = "="""""
    CheckExtent ws.Range("B2"), False, 0, "rows.formulaEmpty"
    ws.Range("B1027").Value2 = CVErr(2042)
    CheckExtent ws.Range("B2"), False, 0, "rows.errorNotBlank"
    ws.Range("B1027").ClearContents
    CheckExtent ws.Range("B2"), False, 0, "rows.empty"
    CheckExtent ws.Range("C2"), False, 0, "rows.blank"
    ws.Range("D4:F4").Value2 = "header"
    CheckExtent ws.Range("D4"), True, 256, "columns"
    CheckExtent ws.Range("D4"), True, 1, "columns.limitOne"
    Set anchor = ws.Cells(ws.Rows.Count, ws.Columns.Count)
    CheckExtent anchor, False, 0, "edge.emptyRow"
    CheckExtent anchor, True, 256, "edge.emptyColumn"
    anchor.Value2 = "last"
    CheckExtent anchor, False, 0, "edge.scalarRow"
    CheckExtent anchor, True, 256, "edge.scalarColumn"
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

Private Sub CheckExtent(ByVal anchor As Object, ByVal columns As Boolean, ByVal limit As Long, ByVal label As String)
    Dim actual As Long, expected As Long
    expected = OldExtent(anchor, columns, limit)
    actual = CLng(Application.Run("'" & ThisWorkbook.Name & "'!modWorkbookCalculation.SnapshotAnchoredExtent", anchor, columns, limit))
    Check label, actual = expected
End Sub

' Независимый прежний oracle: именно первая Trim/CStr-пустая, не UsedRange.
Private Function OldExtent(ByVal anchor As Object, ByVal columns As Boolean, ByVal limit As Long) As Long
    Dim available As Long, i As Long, value As Variant
    If columns Then available = anchor.Worksheet.Columns.Count - anchor.Column + 1 Else available = anchor.Worksheet.Rows.Count - anchor.Row + 1
    If limit > 0 And limit < available Then available = limit
    For i = 0 To available - 1
        If columns Then value = anchor.Offset(0, i).Value2 Else value = anchor.Offset(i, 0).Value2
        If Len(Trim$(CStr(value))) = 0 Then Exit Function
        OldExtent = OldExtent + 1
    Next i
End Function

' ДЛЯ ТЕСТОВ: фиксирует только неуспешные assertions, не меняет expected/tolerance.
Private Sub Check(ByVal label As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & label & vbCrLf
    End If
End Sub

' ДЛЯ ТЕСТОВ: длительность по одному монотонному таймеру, не VBA Timer.
Private Function SecondsSince(ByVal start As Double) As Double
    SecondsSince = PerformanceNow() - start
End Function

' ДЛЯ ТЕСТОВ: монотонный высокоточный таймер; одинаковый код у A и B.
Private Function PerformanceNow() As Double
    Dim counter As Currency, frequency As Currency
    If QueryPerformanceCounter(counter) = 0 Or QueryPerformanceFrequency(frequency) = 0 Then _
        Err.Raise vbObjectError + 4980, "modTestPerformance", "Высокоточный таймер недоступен."
    PerformanceNow = counter / frequency
End Function

' ДЛЯ ТЕСТОВ: сериализует все биты Double, не округляет и не меняет tolerance.
Private Function DoubleBits(ByVal value As Double) As String
    Dim number As TDoubleValue, bytes As TDoubleBytes, i As Long
    number.value = value: LSet bytes = number
    For i = 1 To 8: DoubleBits = DoubleBits & Right$("0" & Hex$(bytes.value(i)), 2): Next i
End Function

' ДЛЯ ТЕСТОВ: инженерные поля solver-а, без изменяемых технических счетчиков.
' Матрица включается только при проверке публичной полной оценки/обычного Solve.
Private Function SolverFingerprint(ByVal solver As CSectionSolver, Optional ByVal includeK As Boolean = True) As String
    Dim values As Variant, value As Variant, i As Long, j As Long
    values = Array(solver.Converged, CLng(solver.FailureCode), solver.StopReason, solver.Epsilon0, solver.KappaX, solver.KappaY, _
        solver.Nint, solver.Mxint, solver.Myint, solver.ResidualN, solver.ResidualMx, solver.ResidualMy, _
        solver.RelativeResidualN, solver.RelativeResidualMx, solver.RelativeResidualMy, _
        solver.MinConcreteStrain, solver.MaxConcreteStrain, solver.MaxConcreteCompressionStress, _
        solver.MinConcreteStress, solver.MaxConcreteStress, solver.HasConcreteCompression, solver.HasConcreteTension, _
        solver.MinSteelStrain, solver.MaxSteelStrain, solver.MaxAbsSteelStrain, solver.MinSteelStress, solver.MaxSteelStress, _
        solver.HasSteelCompression, solver.HasSteelTension, solver.Iterations, solver.LoadStepsCompleted, _
        solver.LastResidualNorm, solver.LastStepNorm, solver.LastMatrixConditionEstimate, solver.MatrixRestartCount)
    For Each value In values
        If VarType(value) = vbDouble Then
            SolverFingerprint = SolverFingerprint & "|" & DoubleBits(CDbl(value))
        Else
            SolverFingerprint = SolverFingerprint & "|" & CStr(value)
        End If
    Next value
    If includeK Then
        For i = 1 To 3: For j = 1 To 3
            SolverFingerprint = SolverFingerprint & "|" & DoubleBits(solver.Tangent(i, j))
        Next j: Next i
    End If
End Function

' ДЛЯ ТЕСТОВ: штатные диаграммы/узлы, две численные ветви и границы одного Solve.
' Frozen baseline не компилирует новые Friend API; численный oracle не меняется.
Public Function RunPerformanceNumericTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
#If PERFORMANCE_CURRENT Then
    Dim section As CSectionModel, concrete As CMaterialDiagram, steel As CMaterialDiagram
    Dim frozenConcrete As CMaterialDiagram, frozenSteel As CMaterialDiagram
    Set section = NumericSection(8)
    NumericMaterials concrete, steel
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
                Check "physical.exact." & CStr(i) & "." & CStr(shift), SolverFingerprint(full, False) = SolverFingerprint(physical, False)
                Check "physical.noK", physical.TangentBuildCount = 0 And full.TangentBuildCount = 1
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
            Check "solve.exact." & CStr(method) & "." & CStr(i), SolverFingerprint(plain) = SolverFingerprint(cached)
            Check "solve.mutableNoReuse", plain.EvaluationReuseCount = 0
            hits = hits + cached.EvaluationReuseCount
        Next i
        Check "solve.reuse." & CStr(method), hits > 0
    Next method
    section.AddConcreteElement 205#, 0#, 100#, 1, "mutation"
    plain.Solve section, concrete, steel, -100000#, 2000000#, -3000000#
    cached.Solve section, frozenConcrete, frozenSteel, -100000#, 2000000#, -3000000#
    Check "solve.changedModel", SolverFingerprint(plain) = SolverFingerprint(cached)
    cached.ClearResult
    Check "solve.clear", cached.EvaluationReuseCount = 0 And cached.TangentBuildCount = 0
#Else
    Check "numeric.currentOnly", False
#End If
    GoTo Finished
Failed:
    Check "numeric.runtime " & CStr(Err.Number) & ": " & Err.Description, False
Finished:
    RunPerformanceNumericTests = mReport & "TOTAL PERFORMANCE NUMERIC: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: детерминированная сетка и четыре стержня; без Excel в ядре.
Private Function NumericSection(ByVal divisions As Long) As CSectionModel
    Dim section As CSectionModel, i As Long, j As Long, step As Double
    Set section = New CSectionModel: step = 400# / divisions
    For i = 1 To divisions: For j = 1 To divisions
        section.AddConcreteElement -200# + (i - 0.5) * step, -200# + (j - 0.5) * step, step * step, _
            1, "cell", vbNullString, "Rectangle", step, step
    Next j: Next i
    For i = -1 To 1 Step 2: For j = -1 To 1 Step 2
        section.AddRebarElement 140# * i, 140# * j, 20#, 0#, "Periodic", 1, "bar"
    Next j: Next i
    Set NumericSection = section
End Function

' ДЛЯ ТЕСТОВ: одинаковые mutable/frozen копии без provider/config-умолчаний.
Private Sub NumericMaterials(ByRef concrete As CMaterialDiagram, ByRef steel As CMaterialDiagram)
    Dim eps(1 To 5) As Double, stress(1 To 5) As Double
    eps(1) = -0.0035: eps(2) = -0.002: eps(3) = 0#: eps(4) = 0.00015: eps(5) = 0.0003
    stress(1) = -30#: stress(2) = -30#: stress(3) = 0#: stress(4) = 2#: stress(5) = 2#
    Set concrete = New CMaterialDiagram: concrete.InitializeFromArrays eps, stress, 5
    eps(1) = -0.02: eps(2) = -0.002: eps(3) = 0#: eps(4) = 0.002: eps(5) = 0.02
    stress(1) = -400#: stress(2) = -400#: stress(3) = 0#: stress(4) = 400#: stress(5) = 400#
    Set steel = New CMaterialDiagram: steel.InitializeFromArrays eps, stress, 5
End Sub

' ДЛЯ ТЕСТОВ: одинаковые 1/10/30 независимых задач; preparation/core отдельно.
' Fingerprint и assertions выполняются после таймера ядра.
Public Function MeasurePerformanceSolver(ByVal method As String, ByVal divisions As Long, ByVal count As Long) As String
    Dim start As Double, prepared As Double, elapsed As Double, section As CSectionModel
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, solvers() As CSectionSolver, i As Long
    start = PerformanceNow()
    Set section = NumericSection(divisions): NumericMaterials concrete, steel
    Set concrete = concrete.Clone: Set steel = steel.Clone
    ReDim solvers(1 To count)
    prepared = SecondsSince(start)
    start = PerformanceNow()
    For i = 1 To count
        Set solvers(i) = New CSectionSolver
        solvers(i).SolverMethod = method: solvers(i).DiagnosticsEnabled = False
        solvers(i).Solve section, concrete, steel, (-1#) ^ i * (100000# + 11000# * i), _
            (i - 15#) * 1000000#, ((7 * i Mod 31) - 15#) * 700000#
    Next i
    elapsed = SecondsSince(start)
    Dim fingerprint As String, evaluations As Long, builds As Long, reuse As Long
    For i = 1 To count
        fingerprint = fingerprint & vbLf & CStr(i) & SolverFingerprint(solvers(i))
        evaluations = evaluations + solvers(i).InternalForceEvaluationCount
#If PERFORMANCE_CURRENT Then
        builds = builds + solvers(i).TangentBuildCount: reuse = reuse + solvers(i).EvaluationReuseCount
#End If
    Next i
    MeasurePerformanceSolver = "seconds=" & NumberText(elapsed) & Chr$(30) & "preparation=" & NumberText(prepared) & _
        Chr$(30) & "elements=" & CStr(section.ConcreteCount) & Chr$(30) & "bars=" & CStr(section.RebarCount) & _
        Chr$(30) & "evaluations=" & CStr(evaluations) & Chr$(30) & "K=" & CStr(builds) & _
        Chr$(30) & "reuse=" & CStr(reuse) & Chr$(30) & "fingerprint=" & fingerprint
End Function

' ДЛЯ ТЕСТОВ: побитовый снимок механических A/I, не времени вычисления.
Private Function PropertiesFingerprint(ByVal props As CSectionPropertiesCalculator) As String
    Dim value As Variant
    For Each value In Array(props.Area, props.StaticMomentX, props.StaticMomentY, props.CentroidX, props.CentroidY, _
            props.Ix, props.Iy, props.Ixy, props.Ixc, props.Iyc, props.Ixyc, props.PrincipalI1, props.PrincipalI2, _
            props.PrincipalAngleRad, props.RadiusX, props.RadiusY, props.PrincipalRadius1, props.PrincipalRadius2)
        PropertiesFingerprint = PropertiesFingerprint & "|" & DoubleBits(CDbl(value))
    Next value
End Function

' ДЛЯ ТЕСТОВ: точный порядок/координаты/дуги и площадь готовой области.
Private Function RegionFingerprint(ByVal region As CConcreteRegion) As String
    If region Is Nothing Then RegionFingerprint = "Nothing": Exit Function
    Dim edges() As TRegionEdge, count As Long, i As Long
    region.CopyEdges edges, count
    RegionFingerprint = region.Source & "|" & CStr(region.SectionRevision) & "|" & CStr(region.LoopCount) & "|" & DoubleBits(region.Area)
    For i = 1 To count
        RegionFingerprint = RegionFingerprint & "|" & CStr(edges(i).LoopID) & DoubleBits(edges(i).X1) & DoubleBits(edges(i).Y1) & _
            DoubleBits(edges(i).X2) & DoubleBits(edges(i).Y2) & DoubleBits(edges(i).CenterX) & DoubleBits(edges(i).CenterY) & DoubleBits(edges(i).Sweep)
    Next i
End Function

' ДЛЯ ТЕСТОВ: общая область и механика не смешивают LC и не удерживаются в
' frozen SP35 snapshot; смены модели, контура и модулей требуют новой подготовки.
Public Function RunPerformancePreparationTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
#If PERFORMANCE_CURRENT Then
    Dim section As CSectionModel, query As CSectionGeometryQuery, fresh As CSectionGeometryQuery
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, first As CConcreteRegion, second As CConcreteRegion
    Set section = NumericSection(8): NumericMaterials concrete, steel
    Set query = New CSectionGeometryQuery: query.Initialize section
    Dim i As Long, j As Long, sharedData As CSP35CrackData, freshData As CSP35CrackData, members As Variant
    For i = 1 To 30
        Set fresh = New CSectionGeometryQuery: fresh.Initialize section
        Set first = query.ConcreteDomain: Set second = fresh.ConcreteDomain
        Check "domain.exact." & CStr(i), RegionFingerprint(first) = RegionFingerprint(second)
        first.SetOwner "changed clone", "SP35", "G1"
        Set second = query.ConcreteDomain
        Check "domain.clone", Len(second.OwnerID) = 0
        Set sharedData = New CSP35CrackData: Set freshData = New CSP35CrackData
        sharedData.Prepare section, 0.0004 + i * 0.000001, 0.0000001, -0.00000015, steel, 10#, 50#, "Max", "6d", False, 0.2, query
        freshData.Prepare section, 0.0004 + i * 0.000001, 0.0000001, -0.00000015, steel, 10#, 50#, "Max", "6d", False
        Check "sp35.count", sharedData.CandidateCount = freshData.CandidateCount And sharedData.GroupCount = freshData.GroupCount
        For j = 1 To sharedData.CandidateCount
            Check "sp35.region", RegionFingerprint(sharedData.CandidateRegion(j)) = RegionFingerprint(freshData.CandidateRegion(j))
            Check "sp35.beta", DoubleBits(sharedData.CandidateBetaDiameterSum(j)) = DoubleBits(freshData.CandidateBetaDiameterSum(j))
            Check "sp35.stress", DoubleBits(sharedData.CandidateSigmaS(j)) = DoubleBits(freshData.CandidateSigmaS(j))
        Next j
        sharedData.Freeze
        members = sharedData.GroupMembers(1)
        Check "sp35.published", members(1) > 0
    Next i
    Check "domain.singleBuild", query.DomainBuildCount = 1
    section.Contours.AddContourCircle "opening", 0#, 0#, 30#, "test", "opening", "Opening"
    query.EnsureSection section
    Set fresh = New CSectionGeometryQuery: fresh.Initialize section
    Set first = query.ConcreteDomain: Set second = fresh.ConcreteDomain
    Check "domain.contoursOnly", RegionFingerprint(first) = RegionFingerprint(second) And query.DomainBuildCount = 2
    Set section = NumericSection(8)
    query.EnsureSection section
    Set fresh = New CSectionGeometryQuery: fresh.Initialize section
    Set first = query.ConcreteDomain: Set second = fresh.ConcreteDomain
    Check "domain.identity", RegionFingerprint(first) = RegionFingerprint(second) And query.DomainBuildCount = 3
    Dim props As CSectionPropertiesCalculator, oracle As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator: Set oracle = New CSectionPropertiesCalculator
    Dim nx As Double, ny As Double, low As Double, high As Double, oldLow As Double, oldHigh As Double
    For i = 1 To 30
        props.EnsureTransformed section, 30000#, 200000#: oracle.CalculateTransformedByModuli section, 30000#, 200000#
        Check "properties.exact", PropertiesFingerprint(props) = PropertiesFingerprint(oracle)
        nx = Cos(CDbl(i)): ny = Sin(CDbl(i))
        For j = 1 To 2
            props.CalculateProjection section, nx, ny, False, low, high
            oracle.CalculateProjection section, nx, ny, False, oldLow, oldHigh
            Check "properties.projection", DoubleBits(low) = DoubleBits(oldLow) And DoubleBits(high) = DoubleBits(oldHigh)
            Check "properties.concreteI", DoubleBits(props.ConcreteInertiaAboutAxis(section, nx, ny, 3#, -2#)) = _
                DoubleBits(oracle.ConcreteInertiaAboutAxis(section, nx, ny, 3#, -2#))
            Check "properties.rebarI", DoubleBits(props.RebarInertiaAboutAxis(section, nx, ny, 3#, -2#)) = _
                DoubleBits(oracle.RebarInertiaAboutAxis(section, nx, ny, 3#, -2#))
        Next j
    Next i
    Check "properties.singleBuild", props.PreparationBuildCount = 1
    props.EnsureTransformed section, 31000#, 200001#: oracle.CalculateTransformedByModuli section, 31000#, 200001#
    Check "properties.changedModuli", PropertiesFingerprint(props) = PropertiesFingerprint(oracle) And props.PreparationBuildCount = 2
    section.AddConcreteElement 210#, 0#, 100#, 1, "mutation"
    props.EnsureTransformed section, 31000#, 200001#: oracle.CalculateTransformedByModuli section, 31000#, 200001#
    Check "properties.changedModel", PropertiesFingerprint(props) = PropertiesFingerprint(oracle) And props.PreparationBuildCount = 3
    props.EnsureConcrete section: oracle.CalculateConcrete section
    Check "properties.changedRepresentation", PropertiesFingerprint(props) = PropertiesFingerprint(oracle) And props.PreparationBuildCount = 4
    props.Clear: props.EnsureConcrete section
    Check "properties.clear", PropertiesFingerprint(props) = PropertiesFingerprint(oracle) And props.PreparationBuildCount = 5
#Else
    Check "preparation.currentOnly", False
#End If
    GoTo Finished
Failed:
    Check "preparation.runtime " & CStr(Err.Number) & ": " & Err.Description, False
Finished:
    RunPerformancePreparationTests = mReport & "TOTAL PERFORMANCE PREPARATION: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: машинная запись времени с точкой для одинакового A/B parser.
Private Function NumberText(ByVal value As Double) As String
    NumberText = Replace$(Format$(value, "0.###############"), ",", ".")
End Function

' ДЛЯ ТЕСТОВ: только частная A/B-копия. 1/10 LC - префиксы 30 разных LC.
' Порядок меняет позиции, но не ID/нагрузки. Адреса читаются по именам/ключам.
Public Sub ConfigurePerformanceFixture(ByVal family As String, ByVal count As Long, _
        ByVal method As String, Optional ByVal shape As String = "Saved", _
        Optional ByVal order As String = "Normal", Optional ByVal extension As String = "Yes")
    If count < 1 Or count > 30 Then Err.Raise 5, "ConfigurePerformanceFixture", "Test count must be 1..30."
    PerformanceSetting "Solver.Method", method
    PerformanceSetting "General.DiagramExtension", extension
    PerformanceSetting "Calculation.ZeroMomentPerDepth", "0"
    If shape <> "Saved" Then
        PerformanceSetting "Geometry.Source", "Generated"
        PerformanceSetting "Geometry.Type", shape
    End If
    Dim capacity As Boolean, crack As Boolean, stability As Boolean
    capacity = InStr(family, "CAPACITY") > 0 Or family = "MIXED"
    crack = InStr(family, "CRACK") > 0 Or family = "MIXED"
    stability = InStr(family, "STABILITY") > 0 Or family = "MIXED"
    If InStr(family, "SP63") > 0 Then
        If crack Then PerformanceSetting "SLS.Crack.Code", "SP63"
        If stability Then PerformanceSetting "Stability.Code", "SP63"
    Else
        If crack Then PerformanceSetting "SLS.Crack.Code", "SP35"
        If stability Then PerformanceSetting "Stability.Code", "SP35"
    End If
    If capacity Then
        If InStr(family, "MULTIPLIER") > 0 Then
            PerformanceSetting "Capacity.SolutionStrategy", "LoadMultiplier"
        Else
            PerformanceSetting "Capacity.SolutionStrategy", "Auto"
        End If
    End If
    Dim profiles As Object, data As Variant, row As Long, col As Long, key As String, headerRow As Long
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    data = profiles.Value2
    For row = 1 To UBound(data, 1)
        If CStr(data(row, 3)) = "PR1" Then headerRow = row: Exit For
    Next row
    If headerRow = 0 Then Err.Raise 5, "ConfigurePerformanceFixture", "Missing profile header."
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
    If loads.Rows.Count < 31 Then Err.Raise 5, "ConfigurePerformanceFixture", "Private fixture needs 30 load rows."
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
    If Left$(family, 9) = "IMPORTED_" Then PreparePerformanceImportedSnapshot units
End Sub

' ДЛЯ ТЕСТОВ: снимок по контракту AutoCADImport с точными дугами и двумя
' отверстиями. Это проверка Results-pipeline, а не новый сеанс настоящего CAD.
' В A и B сетка/контуры/стержни одинаковы; механическая сетка не заменяется контуром.
Private Sub PreparePerformanceImportedSnapshot(ByVal units As CUnitSystem)
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
    PerformanceSetting "Geometry.Source", "AutoCAD"
End Sub

' ДЛЯ ТЕСТОВ: настройка частной копии по фактическому ключу.
Private Sub PerformanceSetting(ByVal key As String, ByVal value As String)
    Dim table As Object, data As Variant, row As Long
    Set table = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    data = table.Value2
    For row = 1 To UBound(data, 1)
        If CStr(data(row, 1)) = key Then table.Cells(row, 2).Value2 = value: Exit Sub
    Next row
    Err.Raise 5, "PerformanceSetting", "Missing setting: " & key
End Sub

' ДЛЯ ТЕСТОВ: Full - настоящий макрос; Staged - те же владельцы с QPC.
' Assertions/fingerprint вне таймера; ни один объект не живет между вызовами.
Public Function MeasurePerformanceBatch(ByVal mode As String) As String
    Dim t As Double, total As Double, preparation As Double, core As Double, clear As Double
    Dim summary As Double, snapshot As Double, plot As Double, packing As Double
    Dim guard As CExcelAppStateGuard, message As String, tech As String, fingerprint As String
    On Error GoTo Failed
    t = PerformanceNow(): total = t
    If mode = "Full" Then
        message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
        total = SecondsSince(total)
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
        preparation = SecondsSince(t)
        t = PerformanceNow(): ClearSectionResultsForWorkbook ThisWorkbook: clear = SecondsSince(t)
        t = PerformanceNow(): batch.Execute: core = SecondsSince(t)
        Dim summaryWriter As CBatchResultWriter, ndm As CNDMResultsWriter, props As CSectionPropertiesCalculator
        Set summaryWriter = New CBatchResultWriter: Set ndm = New CNDMResultsWriter
        t = PerformanceNow(): Set props = PrepareSectionSnapshot(section, provider): packing = SecondsSince(t)
        t = PerformanceNow(): summaryWriter.WriteSummary ThisWorkbook, batch, units, section: summary = SecondsSince(t)
        t = PerformanceNow(): ndm.WriteResults ThisWorkbook, section, props, batch, units: snapshot = SecondsSince(t)
        t = PerformanceNow()
        If settings.GetRequiredBoolean("Plot.Enabled") And settings.GetRequiredBoolean("Plot.AutoUpdateAfterCalculation") Then
            If batch.StateAvailableCount > 0 Then UpdateSectionPlotForWorkbook ThisWorkbook, False
        End If
        plot = SecondsSince(t): total = SecondsSince(total)
        tech = "elements=" & CStr(section.ConcreteCount) & Chr$(30) & "bars=" & CStr(section.RebarCount) & Chr$(30) & _
            "solves=" & CStr(batch.SolverCallCount)
#If PERFORMANCE_CURRENT Then
        tech = tech & Chr$(30) & "domainBuilds=" & CStr(batch.CrackDomainBuildCount) & Chr$(30) & _
            "mechanicalBuilds=" & CStr(batch.MechanicalPreparationBuildCount) & Chr$(30) & "released=" & CStr(batch.RunPreparationReleased)
#End If
        Dim i As Long, result As CCombinationResult
        For i = 1 To batch.Count
            Set result = batch.ResultAt(i)
            fingerprint = fingerprint & batch.CombinationID(i) & "|" & MetaFingerprint(result.OverallMeta) & "|" & _
                MetaFingerprint(result.DirectStateMeta) & "|" & MetaFingerprint(result.CapacityMeta) & "|" & _
                MetaFingerprint(result.CrackFormationMeta) & "|" & MetaFingerprint(result.CrackCurrentStateMeta) & "|" & _
                MetaFingerprint(result.CrackWidthMeta) & "|" & MetaFingerprint(result.LongitudinalCrackMeta) & "|" & _
                MetaFingerprint(result.StabilityMeta) & vbLf
        Next i
        guard.Restore
    End If
    MeasurePerformanceBatch = "seconds=" & NumberText(total) & Chr$(30) & "preparation=" & NumberText(preparation) & Chr$(30) & _
        "core=" & NumberText(core) & Chr$(30) & "clear=" & NumberText(clear) & Chr$(30) & "packing=" & NumberText(packing) & Chr$(30) & _
        "summary=" & NumberText(summary) & Chr$(30) & "snapshot=" & NumberText(snapshot) & Chr$(30) & "plot=" & NumberText(plot) & _
        Chr$(30) & tech & Chr$(30) & "fingerprint=" & fingerprint & PerformanceResultsFingerprint()
    Exit Function
Failed:
    Dim number As Long, description As String
    number = Err.Number: description = Err.Description
    If Not guard Is Nothing Then guard.Restore
    Err.Raise number, "MeasurePerformanceBatch", description
End Function

' ДЛЯ ТЕСТОВ: один и тот же LC должен дать точные данные в пакете, повторном
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
    ConfigurePerformanceFixture family, 30, "Newton"
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    values = loads.Offset(1, 0).Resize(loads.Rows.Count - 1, loads.Columns.Count).Value2
    Set reference = CreateObject("Scripting.Dictionary")
    Set batch = BuildPerformanceTestBatch(section): batch.Execute
    For i = 1 To batch.Count
        reference.Add batch.CombinationID(i), CombinationFingerprint(batch, section, i)
    Next i
    Check "isolation.count", reference.Count = 30
    Dim retained As CCombinationResult, retainedText As String
    Set retained = batch.ResultAt(1): retainedText = ResultFingerprint(retained, section)
    batch.Execute
    For i = 1 To batch.Count
        Check "isolation.repeat." & batch.CombinationID(i), CombinationFingerprint(batch, section, i) = reference(batch.CombinationID(i))
    Next i
    Check "isolation.retained", ResultFingerprint(retained, section) = retainedText
#If PERFORMANCE_CURRENT Then
    Check "isolation.released", batch.RunPreparationReleased
#End If
    ' Ошибка ввода содержит фактический адрес, который меняется при перестановке.
    ' Поэтому MIXED проверяется отдельно с сохранением исходного места каждой строки.
    If family <> "MIXED" Then
        For Each order In Array("Reverse", "Shuffle")
            ConfigurePerformanceFixture family, 30, "Newton", "Saved", CStr(order)
            Set batch = BuildPerformanceTestBatch(section): batch.Execute
            For i = 1 To batch.Count
                Check "isolation." & CStr(order) & "." & batch.CombinationID(i), _
                    CombinationFingerprint(batch, section, i) = reference(batch.CombinationID(i))
            Next i
        Next order
    End If
    For i = 1 To 30
        ReDim singles(1 To UBound(values, 1), 1 To UBound(values, 2))
        For col = 1 To UBound(values, 2): singles(i, col) = values(i, col): Next col
        loads.Offset(1, 0).Resize(UBound(singles, 1), UBound(singles, 2)).Value2 = singles
        Set batch = BuildPerformanceTestBatch(section): batch.Execute
        Check "isolation.single.count." & CStr(i), batch.Count = 1
        Check "isolation.single." & CStr(i), CombinationFingerprint(batch, section, 1) = reference(batch.CombinationID(1))
    Next i
    GoTo Restore
Failed:
    Check "isolation.runtime " & CStr(Err.Number) & ": " & Err.Description, False
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

' ДЛЯ ТЕСТОВ: штатные readers и владельцы без записи Results и нового warm start.
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

' ДЛЯ ТЕСТОВ: рабочие нагрузки и все result-поддеревья одного ID, без таймеров.
Private Function CombinationFingerprint(ByVal batch As CBatchSectionCalculator, ByVal section As CSectionModel, ByVal index As Long) As String
    CombinationFingerprint = DoubleBits(batch.N(index)) & DoubleBits(batch.Mx(index)) & DoubleBits(batch.My(index)) & _
        DoubleBits(batch.UserMx(index)) & DoubleBits(batch.UserMy(index)) & ResultFingerprint(batch.ResultAt(index), section)
End Function

' ДЛЯ ТЕСТОВ: точные поля ширины, несущей, устойчивости и каждого named-state.
' Технические времена/счетчики не инженерные expected; их проверяют замеры.
Private Function ResultFingerprint(ByVal result As CCombinationResult, ByVal section As CSectionModel) As String
    Dim text As String, state As CSectionStateResult, i As Long
    text = MetaFingerprint(result.OverallMeta) & MetaFingerprint(result.DirectStateMeta) & MetaFingerprint(result.CapacityMeta) & _
        MetaFingerprint(result.CrackFormationMeta) & MetaFingerprint(result.CrackCurrentStateMeta) & MetaFingerprint(result.CrackWidthMeta) & _
        MetaFingerprint(result.LongitudinalCrackMeta) & MetaFingerprint(result.StabilityMeta)
    text = text & ScalarFieldsFingerprint(result, "Status NormalCrackStatus CrackSummaryStatus ExtensionUsed GoverningUtil ProfileId DiagramExtensionEnabled")
    text = text & ScalarFieldsFingerprint(result.StrengthResult.DirectState, "Status StateAvailable StateConverged ExtensionUsed ExtensionEnabled StrainReserve")
    text = text & ScalarFieldsFingerprint(result.StrengthResult.Capacity, "Status LimitState SolutionMethod PathResolved LambdaCapacity UtilCapacity NUltimate MxUltimate MyUltimate CapacityStateN CapacityStateMx CapacityStateMy MomentUltimate CapacityStateAvailable CapacityEpsilon0 CapacityKappaX CapacityKappaY PureAxialForMethod ReserveFactor SearchExecuted")
    text = text & ScalarFieldsFingerprint(result.CrackResult.Formation, "CrackFormed ConfirmedCracked ConfirmedNotCracked FormationMethod LambdaCrc Ncrc FormationNcrc Mcrc Ared RbtSer StopReason HasLimitPoint FallbackPsi1 CentralTensionBranch Converged")
    text = text & ScalarFieldsFingerprint(result.CrackResult.Width, "CrackWidth AllowableCrackWidth Utilization Phi1 Phi2 Phi3 PsiS SigmaS SigmaSCrc SteelEs AsTension Abt DsEquivalent CrackSpacingRaw CrackSpacing TensionRebarCount TensionRebarIds SectionDepthH CoverA TensionDepth EffectiveZoneDepth CentralTensionBranch StopReason StandardCode SP35CriticalCandidate SP35ReinforcementRadius SP35PsiCm CriticalAnchorGroupID ReserveFactor")
    text = text & ScalarFieldsFingerprint(result.CrackResult.Longitudinal, "Status MaxCompressionStress RbMc2 Utilization ReserveFactor")
    text = text & ScalarFieldsFingerprint(result.StabilityResult, "Status SummaryReserve Code Branch ReferenceX ReferenceY DesignN DesignMx DesignMy SustainedN Moment1 Moment2 SustainedMoment1 SustainedMoment2 DesignMoment1 DesignMoment2 CriticalForce ReserveFactor Utilization Ncr1 Ncr2 Nultimate1 Nultimate2 Reserve1 Reserve2 Eta1 Eta2 PhiL1 PhiL2 PhiLTable1 PhiLTable2 PhiM1 PhiM2 PhiValue1 PhiValue2 PhiP1 PhiP2 Eccentricity1 Eccentricity2 AccidentalEcc1 AccidentalEcc2 EffectiveLength1 EffectiveLength2 Slenderness1 Slenderness2 Depth1 Depth2 Radius1 Radius2 CoreDistance1 CoreDistance2 Delta1 Delta2 Kb1 Kb2 Ks1 Ks2 StiffnessD1 StiffnessD2 PlaneBranch1 PlaneBranch2 PlaneApplicable1 PlaneApplicable2 PlanePassed1 PlanePassed2")
    For i = 1 To result.StateRepository.StateCount
        Set state = result.StateRepository.StateAt(i)
        text = text & ScalarFieldsFingerprint(state, "StateType StateTypeText MaterialModelRole MaterialModelRoleText Epsilon0 KappaX KappaY ExtensionUsed WithinPhysicalRange Converged StopReason Status InternalStatus ResultCode ResultComment TargetN TargetMx TargetMy Nint Mxint Myint ResidualN ResidualMx ResidualMy RelativeResidualN RelativeResidualMx RelativeResidualMy LastResidualNorm IterationCount MinConcreteStrain MaxConcreteStrain MinSteelStrain MaxSteelStrain MinConcreteStress MaxConcreteStress MinSteelStress MaxSteelStress HasConcreteCompression HasConcreteTension HasSteelCompression HasSteelTension")
        text = text & state.MaterialSpec.SpecKey & ArrayFingerprint(state.PrepareElementSnapshot(section))
    Next i
    Dim width As CCrackWidthResult, data As CSP35CrackData, j As Long
    Set width = result.CrackResult.Width
    If Not width Is Nothing Then
        text = text & RegionFingerprint(width.InteractionRegion)
        Set data = width.SP35Data
        If Not data Is Nothing Then
            text = text & ScalarFieldsFingerprint(data, "CandidateCount GroupCount NormalX NormalY NeutralProjection GroupGapTolerance RowTolerance NeighborRatioLimit RadiusDiameterMode InteractionRadiusMode CentralTension SectionRevision")
            For j = 1 To data.GroupCount
                text = text & DoubleBits(data.GroupX(j)) & DoubleBits(data.GroupY(j)) & DoubleBits(data.GroupArea(j)) & _
                    DoubleBits(data.GroupBeta(j)) & CStr(data.GroupRow(j)) & ArrayFingerprint(data.GroupMembers(j))
            Next j
            For j = 1 To data.CandidateCount
                text = text & CStr(data.CandidateAvailable(j))
                If Not data.CandidateAvailable(j) Then GoTo NextIsolationCandidate
                text = text & RegionFingerprint(data.CandidateRegion(j)) & ArrayFingerprint(data.CandidateRows(j)) & _
                    CStr(data.CandidateAnchorGroup(j)) & CStr(data.CandidateStressBar(j)) & DoubleBits(data.CandidateSigmaS(j)) & _
                    DoubleBits(data.CandidateBetaDiameterSum(j)) & DoubleBits(data.CandidateSideDiameter(j)) & DoubleBits(data.CandidateRowDiameter(j)) & _
                    DoubleBits(data.CandidateLeftDistance(j)) & DoubleBits(data.CandidateRightDistance(j)) & DoubleBits(data.CandidateNormalRadius(j)) & _
                    CStr(data.CandidateReferenceRow(j)) & CStr(data.CandidateRowCount(j)) & DoubleBits(data.CandidateCrackWidth(j))
NextIsolationCandidate:
            Next j
        End If
    End If
    ResultFingerprint = text
End Function

' ДЛЯ ТЕСТОВ: только явно перечисленные скалярные getters; нет округления Double.
Private Function ScalarFieldsFingerprint(ByVal value As Object, ByVal fields As String) As String
    If value Is Nothing Then ScalarFieldsFingerprint = "<Nothing>": Exit Function
    Dim name As Variant, item As Variant
    For Each name In Split(fields, " ")
        item = CallByName(value, CStr(name), VbGet)
        ScalarFieldsFingerprint = ScalarFieldsFingerprint & CStr(name) & "=" & CellFingerprint(item) & "|"
    Next name
End Function

' ДЛЯ ТЕСТОВ: сохраняет размерность и точные значения массива одного состояния.
Private Function ArrayFingerprint(ByVal data As Variant) As String
    If Not IsArray(data) Then ArrayFingerprint = CellFingerprint(data): Exit Function
    Dim dimensions As Long, row As Long, col As Long, item As Variant
    On Error Resume Next
    dimensions = UBound(data, 2)
    On Error GoTo 0
    If dimensions = 0 Then
        ArrayFingerprint = "1D|" & CStr(LBound(data)) & "|" & CStr(UBound(data))
        For Each item In data: ArrayFingerprint = ArrayFingerprint & CellFingerprint(item) & "|": Next item
    Else
        ArrayFingerprint = "2D|" & CStr(UBound(data, 1)) & "|" & CStr(UBound(data, 2))
        For row = LBound(data, 1) To UBound(data, 1)
            For col = LBound(data, 2) To UBound(data, 2)
                ArrayFingerprint = ArrayFingerprint & CellFingerprint(data(row, col)) & "|"
            Next col
        Next row
    End If
End Function

' ДЛЯ ТЕСТОВ: статус, код, применимость и комментарий без подмены внешним текстом.
Private Function MetaFingerprint(ByVal meta As CResultMeta) As String
    MetaFingerprint = CStr(meta.InternalStatus) & ":" & CStr(meta.ResultCode) & ":" & CStr(meta.Applies) & ":" & _
        CStr(meta.Calculated) & ":" & meta.ResultComment
End Function

' ДЛЯ ТЕСТОВ: числа как биты Double, остальные значения - точный текст/тип.
' Исключены только RunID и summary solves/time. RunID проверяется внутри снимка.
Public Function PerformanceResultsFingerprint() As String
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
            rows = OldExtent(anchor, False, 0): columns = ndm.OutputColumnCount(CStr(name))
            If rows < 1 Then rows = 1
        End If
        data = anchor.Resize(rows, columns).Value2
        ReDim lines(0 To rows)
        lines(0) = CStr(name) & "|" & CStr(rows) & "|" & CStr(columns)
        For row = 1 To rows
            line = vbNullString
            For col = 1 To columns
                value = data(row, col)
                If Left$(CStr(name), 6) = "rngNDM" And col = 1 And row > 1 Then
                    If Len(runID) = 0 Then runID = CStr(value)
                    If CStr(value) <> runID Then Err.Raise 5, "PerformanceResultsFingerprint", "Inconsistent RunID."
                    value = "<RunID>"
                End If
                If name = "rngBatchSummary" And col = 1 And (row = 3 Or row = 4) Then value = "<Technical>"
                line = line & CellFingerprint(value) & "|"
            Next col
            lines(row) = line
        Next row
        result = result & Join(lines, vbLf) & vbLf
    Next name
    PerformanceResultsFingerprint = result
End Function

' ДЛЯ ТЕСТОВ: различает Empty/error/типы и сохраняет биты числовой ячейки.
Private Function CellFingerprint(ByVal value As Variant) As String
    If IsError(value) Then
        CellFingerprint = "E:" & CStr(value)
    ElseIf IsEmpty(value) Then
        CellFingerprint = "Empty"
    ElseIf VarType(value) = vbDouble Then
        CellFingerprint = "D:" & DoubleBits(CDbl(value))
    Else
        CellFingerprint = "T:" & CStr(VarType(value)) & ":" & CStr(Len(CStr(value))) & ":" & CStr(value)
    End If
End Function

' ДЛЯ ТЕСТОВ: повторяет настоящий passive export-reader без запуска AutoCAD.
' Геометрия/цикл выбора LC остаются включены; assertions вне замера.
Public Function MeasurePerformanceExport(ByVal count As Long) As String
    Dim t As Double, i As Long, elapsed As Double, values() As String
    ReDim values(1 To count)
    t = PerformanceNow()
    For i = 1 To count
        values(i) = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    Next i
    elapsed = SecondsSince(t)
    For i = 2 To count
        If values(i) <> values(1) Then Err.Raise 5, "MeasurePerformanceExport", "Read changed within measurement."
    Next i
    MeasurePerformanceExport = "seconds=" & NumberText(elapsed) & Chr$(30) & "fingerprint=" & values(1) & vbLf & PerformanceResultsFingerprint()
End Function

' ДЛЯ ТЕСТОВ: следующая операция видит новые данные; ошибка не удерживает кэш.
Public Function RunPerformanceExportCacheTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    Dim table As Object, data As Variant, row As Long, original As String, changed As String
    Dim oldValue As Variant, oldHeader As Variant, errorNumber As Long, restored As Boolean
    On Error GoTo Failed
    Set table = ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange
    Set table = table.Resize(OldExtent(table.Cells(1, 1), False, 0), 9)
    data = table.Value2
    original = Audit02ReadExportSnapshotForTests(ThisWorkbook)
#If PERFORMANCE_CURRENT Then
    Check "export.released", InStr(PerformanceExportReadDiagnosticsForTests(), "released=True") > 0
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
    Check "export.fresh.operation", changed <> original
    table.Cells(row, 8).Formula = oldValue: restored = True
    Check "export.restored", Audit02ReadExportSnapshotForTests(ThisWorkbook) = original
    oldHeader = table.Cells(1, 8).Formula
    table.Cells(1, 8).Value2 = "broken"
    On Error Resume Next
    changed = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    errorNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "export.corrupt.rejected", errorNumber <> 0
#If PERFORMANCE_CURRENT Then
    Check "export.failure.released", InStr(PerformanceExportReadDiagnosticsForTests(), "released=True") > 0
#End If
    table.Cells(1, 8).Formula = oldHeader
    Check "export.after.failure", Audit02ReadExportSnapshotForTests(ThisWorkbook) = original
    GoTo Restore
Failed:
    Check "export.runtime " & CStr(Err.Number) & ": " & Err.Description, False
Restore:
    On Error Resume Next
    If Not restored And row > 1 Then table.Cells(row, 8).Formula = oldValue
    If Not IsEmpty(oldHeader) Then table.Cells(1, 8).Formula = oldHeader
    On Error GoTo 0
    RunPerformanceExportCacheTests = mReport & "TOTAL PERFORMANCE EXPORT CACHE: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed)
End Function

' ДЛЯ ТЕСТОВ: отделяет B04 от дорогого solve, а P05 от записи чисел.
' Фактическая стоимость 30 вызовов оценивается по серии 10000, не по нулевому Timer.
Public Function MeasurePerformanceOverhead() As String
    Dim catalog As CCalculationProfileCatalog, profile As CCalculationProfile, meta As CResultMeta, copy As CResultMeta
    Dim i As Long, t As Double, validation As Double, cloning As Double, formatting As Double, errorText As String
    Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
    Set profile = catalog.ProfileById("PR1")
    Set meta = New CResultMeta: meta.SetResult rsSuccess, rcCheckPassed, rkDirectState, "Расчет выполнен."
    t = PerformanceNow()
    For i = 1 To 10000
        If Not profile.ValidateForCalculation(errorText) Then Err.Raise 5, "MeasurePerformanceOverhead", errorText
    Next i
    validation = SecondsSince(t)
    t = PerformanceNow()
    For i = 1 To 10000: Set copy = meta.Clone: Next i
    cloning = SecondsSince(t)
    Dim names As Variant, name As Variant, ndm As CNDMResultsWriter, anchor As Object, rows As Long, columns As Long, title As String
    Set ndm = New CNDMResultsWriter
    names = Array("rngNDMSectionGeometry", "rngNDMSectionContours", "rngNDMElementResults", "rngNDMSectionProperties", _
        "rngNDMMaterialDiagrams", "rngNDMSectionAnnotations")
    For Each name In names
        Set anchor = ThisWorkbook.Names.Item(CStr(name)).RefersToRange.Cells(1, 1)
        rows = OldExtent(anchor, False, 0): If rows < 1 Then rows = 1
        columns = ndm.OutputColumnCount(CStr(name)): title = CStr(anchor.Offset(-1, 0).Value2)
        t = PerformanceNow(): ndm.FormatResultsBlock anchor, rows, columns, title
        formatting = formatting + SecondsSince(t)
    Next name
    MeasurePerformanceOverhead = "validation10000=" & NumberText(validation) & Chr$(30) & "clone10000=" & NumberText(cloning) & _
        Chr$(30) & "snapshotFormat=" & NumberText(formatting)
End Function
