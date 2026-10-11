Attribute VB_Name = "modPlateauSlopeProbe"
Option Explicit

' ДЛЯ ТЕСТОВ: согласованный наклон плато и счетчики только экспериментальной книги.
' Рабочая книга не содержит этот модуль и не читает его параметры.
Public Alpha As Double
Public SimpleInitialGuess As Boolean
Public SolverCalls As Long, StateIterations As Long, UltimateRuns As Long, UltimateIterations As Long
Public ActiveSetIterations As Long, DirectGuesses As Long, ExtensionGuesses As Long, RetryCalls As Long
Public CapacityRetries As Long, MultiplierFallbacks As Long, FormationFullStarts As Long
Public WarmRestarts As Long, ProjectedRecoveries As Long, StateRegularizedSteps As Long
Private mRecords As String

' Меняет только ненулевые горизонтальные участки физической диаграммы.
' Начало плато сохраняет исходное напряжение. Сжатая ветвь обходится от нуля
' к отрицательным деформациям; растянутая - от нуля к положительным.
Public Sub ApplySlope(ByRef strains() As Double, ByRef stresses() As Double, ByVal count As Long, _
        ByVal compressionE As Double, ByVal tensionE As Double)
    If Alpha = 0# Then Exit Sub
    If Alpha < 0# Or compressionE <= 0# Or tensionE <= 0# Then Err.Raise 5, , "Invalid plateau slope parameters."
    Dim zero As Long, i As Long, previous As Double, original As Double, offset As Double
    For i = 1 To count
        If strains(i) = 0# And stresses(i) = 0# Then zero = i: Exit For
    Next i
    If zero = 0 Then Err.Raise 5, , "Diagram has no origin."
    previous = 0#: offset = 0#
    For i = zero - 1 To 1 Step -1
        original = stresses(i)
        If original <> 0# And original = previous Then _
            offset = offset + Alpha * compressionE * (strains(i) - strains(i + 1))
        stresses(i) = original + offset
        previous = original
    Next i
    previous = 0#: offset = 0#
    For i = zero + 1 To count
        original = stresses(i)
        If original <> 0# And original = previous Then _
            offset = offset + Alpha * tensionE * (strains(i) - strains(i - 1))
        stresses(i) = original + offset
        previous = original
    Next i
End Sub

Public Sub ResetCounters()
    SolverCalls = 0: StateIterations = 0: UltimateRuns = 0: UltimateIterations = 0
    ActiveSetIterations = 0: DirectGuesses = 0: ExtensionGuesses = 0: RetryCalls = 0
    CapacityRetries = 0: MultiplierFallbacks = 0: FormationFullStarts = 0
    WarmRestarts = 0: ProjectedRecoveries = 0: StateRegularizedSteps = 0
End Sub

Public Sub BeginRun(ByVal alphaValue As Double, ByVal simple As Boolean)
    Alpha = alphaValue: SimpleInitialGuess = simple
    mRecords = vbNullString
    ResetCounters
End Sub

Public Function Records() As String
    Records = mRecords
End Function

' Полное время штатного расчета; счетчики снимаются до независимой проверки снимков.
Public Sub CaptureResult(ByVal id As String, ByVal elapsed As Double, ByVal result As CCombinationResult)
    Dim cap As CCapacityResult, formation As CCrackFormationResult, width As CCrackWidthResult
    Dim i As Long, state As CSectionStateResult, failedStates As Long, failedResults As Long, meta As Variant
    Set cap = result.StrengthResult.Capacity
    Set formation = result.CrackResult.Formation
    Set width = result.CrackResult.Width
    For i = 1 To result.StateRepository.StateCount
        Set state = result.StateRepository.StateAt(i)
        If state.InternalStatus = rsNumericalFailure Then failedStates = failedStates + 1
        mRecords = mRecords & "PLATEAU_STATE|id=" & id & "|type=" & state.StateTypeText & _
            "|converged=" & CStr(state.Converged) & "|status=" & state.Status & _
            "|physical=" & CStr(state.WithinPhysicalRange) & "|extension=" & CStr(state.ExtensionUsed) & _
            "|N=" & Number(state.Nint) & "|Mx=" & Number(state.Mxint) & "|My=" & Number(state.Myint) & _
            "|eps0=" & Number(state.Epsilon0) & "|kx=" & Number(state.KappaX) & "|ky=" & Number(state.KappaY) & _
            "|minConcrete=" & Number(state.MinConcreteStrain) & "|maxConcrete=" & Number(state.MaxConcreteStrain) & _
            "|minSteel=" & Number(state.MinSteelStrain) & "|maxSteel=" & Number(state.MaxSteelStrain) & vbCrLf
    Next i
    For Each meta In Array(result.DirectStateMeta, result.CapacityMeta, result.CrackFormationMeta, _
            result.CrackCurrentStateMeta, result.CrackWidthMeta, result.LongitudinalCrackMeta)
        If meta.InternalStatus = rsNumericalFailure Then failedResults = failedResults + 1
    Next meta
    mRecords = mRecords & "PLATEAU_CASE|id=" & id & "|alpha=" & Number(Alpha) & _
        "|simple=" & CStr(SimpleInitialGuess) & "|seconds=" & Number(elapsed) & _
        "|solverCalls=" & CStr(SolverCalls) & "|stateIterations=" & CStr(StateIterations) & _
        "|ultimateRuns=" & CStr(UltimateRuns) & "|ultimateIterations=" & CStr(UltimateIterations) & _
        "|activeSetIterations=" & CStr(ActiveSetIterations) & "|directGuesses=" & CStr(DirectGuesses) & _
        "|extensionGuesses=" & CStr(ExtensionGuesses) & "|retryCalls=" & CStr(RetryCalls) & _
        "|capacityRetries=" & CStr(CapacityRetries) & "|multiplierFallbacks=" & CStr(MultiplierFallbacks) & _
        "|formationFullStarts=" & CStr(FormationFullStarts) & "|warmRestarts=" & CStr(WarmRestarts) & _
        "|projectedRecoveries=" & CStr(ProjectedRecoveries) & "|stateRegularizedSteps=" & CStr(StateRegularizedSteps) & _
        "|numFailStates=" & CStr(failedStates) & "|numFailResults=" & CStr(failedResults) & _
        "|overall=" & result.Status & "|capacityStatus=" & cap.Status & _
        "|capacityMethod=" & cap.SolutionMethod & "|capacityPath=" & cap.PathResolved & _
        "|lambdaUlt=" & Number(cap.LambdaCapacity) & "|Nult=" & Number(cap.NUltimate) & _
        "|MxUlt=" & Number(cap.MxUltimate) & "|MyUlt=" & Number(cap.MyUltimate) & _
        "|formationStatus=" & CStr(formation.ResultMeta.InternalStatus) & "|formationMethod=" & formation.FormationMethod & _
        "|hasMcrc=" & CStr(formation.HasLimitPoint) & "|lambdaCrc=" & Number(formation.LambdaCrc) & _
        "|Mcrc=" & Number(formation.Mcrc) & "|Ncrc=" & Number(formation.FormationNcrc) & _
        "|widthStatus=" & CStr(width.ResultMeta.InternalStatus) & "|width=" & Number(width.CrackWidth) & _
        "|sigmaS=" & Number(width.SigmaS) & "|Ar=" & Number(width.Abt) & _
        "|As=" & Number(width.AsTension) & "|anchor=" & width.CriticalAnchorGroupID & vbCrLf
End Sub

Public Function Number(ByVal value As Double) As String
    Number = Replace$(Format$(value, "0.0000000000000000E+00"), ",", ".")
End Function

' ДЛЯ ТЕСТОВ: проверяет реальные диаграммы, производные, Clone и extension.
Public Function RunPlateauMaterialChecks() As String
    On Error GoTo Failed
    Dim concrete As New CConcreteMaterialParameters, steel As New CSteelMaterialParameters
    Dim provider As CMaterialModelProvider, reference As CMaterialModelProvider, spec As CMaterialModelSpec
    Dim a As Variant, valueSet As Variant, kind As Variant, tension As Variant, material As CMaterialDiagram
    Dim baseline As CMaterialDiagram, originalAlpha As Double, output As String, passed As Long, failures As Long
    originalAlpha = Alpha
    concrete.Initialize 14.5, 1.05, 18.5, 1.6, 30000#, 27000#, rbMc2:=22#
    steel.Initialize 310#, 350#, 360#, 400#, 205000#, 200000#, 0.018, 0.025, 0.014, 0.015
    For Each a In Array(0#, 0.00001, 0.0001, 0.001)
        Alpha = 0#: Set reference = New CMaterialModelProvider
        reference.InitializeFromParameters concrete, steel, diagramExtensionEnabled:=True
        Alpha = CDbl(a): Set provider = New CMaterialModelProvider
        provider.InitializeFromParameters concrete, steel, diagramExtensionEnabled:=True
        For Each valueSet In Array("ULS(I)", "SLS(II)")
            For Each kind In Array("TwoLine", "ThreeLine")
                For Each tension In Array("Ignore", "UseDiagram")
                    Set spec = New CMaterialModelSpec: spec.Initialize CStr(valueSet), CStr(kind), CStr(tension), CStr(kind)
                    Set material = provider.ConcreteMaterialFromSpec(spec)
                    ' Reference provider creates profile diagrams lazily: keep alpha zero during that call.
                    Alpha = 0#: Set baseline = reference.ConcreteMaterialFromSpec(spec): Alpha = CDbl(a)
                    CheckDiagram material, baseline, concrete.Eb, concrete.Ebt, CDbl(a), output, passed, failures
                    If CStr(tension) = "Ignore" Then
                        CheckValue material.GetStress(0.01) = 0# And material.GetTangentModulus(0.01) = 0#, _
                            "ignored concrete tension remains zero", output, passed, failures
                    End If
                    CheckExtension provider.ConcreteMaterialForEquilibriumFromSpec(spec), material, concrete.Eb, concrete.Ebt, output, passed, failures
                Next tension
                Set material = provider.SteelMaterialFromSpec(spec)
                Alpha = 0#: Set baseline = reference.SteelMaterialFromSpec(spec): Alpha = CDbl(a)
                CheckDiagram material, baseline, steel.Esc, steel.Es, CDbl(a), output, passed, failures
                CheckExtension provider.SteelMaterialForEquilibriumFromSpec(spec), material, steel.Esc, steel.Es, output, passed, failures
            Next kind
        Next valueSet
    Next a
    Alpha = originalAlpha
    RunPlateauMaterialChecks = output & "TOTAL_PLATEAU_MATERIAL: passed=" & CStr(passed) & "; failed=" & CStr(failures) & vbCrLf
    Exit Function
Failed:
    Alpha = originalAlpha
    RunPlateauMaterialChecks = output & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
End Function

Private Sub CheckDiagram(ByVal material As CMaterialDiagram, ByVal baseline As CMaterialDiagram, _
        ByVal compressionE As Double, ByVal tensionE As Double, ByVal alphaValue As Double, _
        ByRef output As String, ByRef passed As Long, ByRef failures As Long)
    Dim i As Long, strain As Double, slope As Double, expected As Double, h As Double, derivative As Double
    Dim clone As CMaterialDiagram, stress As Double, tangent As Double
    CheckValue material.PhysicalCompressionStrain = baseline.PhysicalCompressionStrain And _
        material.PhysicalTensionStrain = baseline.PhysicalTensionStrain, "unchanged physical limits", output, passed, failures
    CheckValue material.PointCount = baseline.PointCount, "unchanged node count", output, passed, failures
    Set clone = material.Clone(True)
    For i = 1 To material.PointCount - 1
        strain = (material.PointStrain(i) + material.PointStrain(i + 1)) / 2#
        expected = baseline.GetTangentModulus(strain)
        If expected = 0# And baseline.GetStress(strain) <> 0# Then
            If strain < 0# Then expected = alphaValue * compressionE Else expected = alphaValue * tensionE
        End If
        slope = material.GetTangentModulus(strain)
        CheckValue Abs(slope - expected) <= 0.00000001 * (1# + Abs(expected)), "plateau slope / unchanged other slopes", output, passed, failures
        h = (material.PointStrain(i + 1) - material.PointStrain(i)) * 0.0001
        derivative = (material.GetStress(strain + h) - material.GetStress(strain - h)) / (2# * h)
        CheckValue Abs(derivative - slope) <= 0.00001 * (1# + Abs(slope)), "Stress derivative equals Tangent", output, passed, failures
        material.EvaluateAtStrain strain, stress, tangent
        CheckValue stress = material.GetStress(strain) And tangent = slope, "combined material API agrees", output, passed, failures
        CheckValue clone.GetStress(strain) = stress And clone.GetTangentModulus(strain) = tangent, "snapshot preserves slope", output, passed, failures
        If alphaValue = 0# Then CheckValue stress = baseline.GetStress(strain), "OFF exact stress", output, passed, failures
    Next i
    ' Saturation outside the defined range is unchanged; no hidden extension is added.
    strain = 2# * material.PhysicalCompressionStrain
    CheckValue material.GetTangentModulus(strain) = 0#, "no implicit compression extension", output, passed, failures
End Sub

Private Sub CheckExtension(ByVal extended As CMaterialDiagram, ByVal physical As CMaterialDiagram, _
        ByVal compressionE As Double, ByVal tensionE As Double, ByRef output As String, ByRef passed As Long, ByRef failures As Long)
    Dim strain As Double, expected As Double, clone As CMaterialDiagram
    strain = physical.PhysicalCompressionStrain * 1.02
    expected = physical.GetStress(physical.PhysicalCompressionStrain) + _
        0.01 * compressionE * (strain - physical.PhysicalCompressionStrain)
    CheckValue Abs(extended.GetStress(strain) - expected) <= 0.00000001 And _
        Abs(extended.GetTangentModulus(strain) - 0.01 * compressionE) <= 0.00000001, "existing compression extension remains consistent", output, passed, failures
    If physical.HasPhysicalTensionLimit Then
        strain = physical.PhysicalTensionStrain * 1.02
        expected = physical.GetStress(physical.PhysicalTensionStrain) + _
            0.01 * tensionE * (strain - physical.PhysicalTensionStrain)
        CheckValue Abs(extended.GetStress(strain) - expected) <= 0.00000001 And _
            Abs(extended.GetTangentModulus(strain) - 0.01 * tensionE) <= 0.00000001, "existing tension extension remains consistent", output, passed, failures
    End If
    Set clone = extended.Clone(True)
    CheckValue clone.PointCount = physical.PointCount And _
        clone.PointStress(1) = physical.PointStress(1), "physical snapshot excludes only extension", output, passed, failures
End Sub

Private Sub CheckValue(ByVal condition As Boolean, ByVal name As String, ByRef output As String, _
        ByRef passed As Long, ByRef failures As Long)
    If condition Then
        passed = passed + 1
    Else
        failures = failures + 1: output = output & "FAIL: " & name & "; alpha=" & Number(Alpha) & vbCrLf
    End If
End Sub

' ДЛЯ ТЕСТОВ: только этот контроль отключает штатное осевое перенаправление.
' Сравнивает именно UltimateStrain с LoadMultiplier, а не название настройки.
Public Function RunPlateauForcedAxial() As String
    On Error GoTo Failed
    Dim settings As New CSystemSettingsReader, units As New CUnitSystem, section As CSectionModel
    Dim provider As CMaterialModelProvider, spec As New CMaterialModelSpec, cap As CCapacitySolver
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, solver As CSectionSolver
    Dim a As Variant, n As Variant, method As Variant, cx As Double, cy As Double, area As Double
    Dim started As Double, elapsed As Double, output As String, completed As Long, originalAlpha As Double
    originalAlpha = Alpha
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    section.ConcreteMeshCentroid cx, cy, area
    spec.Initialize "ULS(I)", "TwoLine", "Ignore", "TwoLine"
    For Each a In Array(0#, 0.0001)
        Alpha = CDbl(a)
        Set provider = New CMaterialModelProvider: provider.Initialize settings, units
        Set concrete = provider.ConcreteMaterialForEquilibriumFromSpec(spec)
        Set steel = provider.SteelMaterialForEquilibriumFromSpec(spec)
        For Each n In Array(-500# * 9806.65, 500# * 9806.65)
            For Each method In Array("UltimateStrain", "LoadMultiplier")
                Set cap = New CCapacitySolver: cap.ApplySettings settings, units
                cap.ConcreteCompressionLimit = concrete.UltimateCompressionStrain
                cap.ConcreteTensionLimitEnabled = False
                cap.SteelCompressionLimit = Abs(steel.UltimateCompressionStrain)
                cap.SteelTensionLimit = Abs(steel.UltimateTensionStrain)
                Call ResetCounters
                started = Timer
                If CStr(method) = "UltimateStrain" Then
                    cap.SolveByUltimateLoadPath section, concrete, steel, 0#, CDbl(n), _
                        0#, CDbl(n) * cy, 0#, CDbl(n) * cx, False
                Else
                    cap.SolveByLoadPathMultiplier section, concrete, steel, 0#, CDbl(n), _
                        0#, CDbl(n) * cy, 0#, CDbl(n) * cx, True
                End If
                elapsed = Timer - started: If elapsed < 0# Then elapsed = elapsed + 86400#
                Set solver = cap.LastSolver
                output = output & "PLATEAU_FORCED_AXIAL|alpha=" & Number(Alpha) & "|N=" & Number(CDbl(n)) & _
                    "|method=" & CStr(method) & "|converged=" & CStr(cap.Converged) & "|lambda=" & Number(cap.LambdaUltimate) & _
                    "|Nult=" & Number(cap.NUltimate) & "|seconds=" & Number(elapsed) & _
                    "|solverCalls=" & CStr(SolverCalls) & "|iterations=" & CStr(StateIterations + UltimateIterations) & _
                    "|activeSetIterations=" & CStr(ActiveSetIterations) & "|reason=" & cap.StopReason & vbCrLf
                If cap.Converged And Not solver Is Nothing Then
                    output = output & "PLATEAU_FORCED_AXIAL_EQUILIBRIUM|confirmed=" & _
                        CStr(solver.ConfirmEquilibrium(cap.NUltimate, cap.MxUltimate, cap.MyUltimate)) & vbCrLf
                End If
                completed = completed + 1
            Next method
        Next n
    Next a
    Alpha = originalAlpha
    RunPlateauForcedAxial = output & "TOTAL_PLATEAU_FORCED_AXIAL: cases=" & CStr(completed) & vbCrLf
    Exit Function
Failed:
    Alpha = originalAlpha
    RunPlateauForcedAxial = output & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function
