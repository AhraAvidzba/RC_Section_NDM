' ДЛЯ ТЕСТОВ: сопоставляет сохраненную и фактическую диаграммы без правки ядра.
Public Function RunFocusedSnapshotDiagnostic() As String
    On Error GoTo Failed
    Dim output As String, settings As CSystemSettingsReader, units As CUnitSystem
    Dim provider As CMaterialModelProvider, state As CSectionStateResult
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, probe As CSectionSolver
    Dim i As Long, strain As Double, actual As Double, saved As Double, item As Variant
    Dim cx As Double, cy As Double, area As Double, centerStrain As Double, tangent As Double, savedTangent As Double
    Dim loadRange As Object, values() As Variant, section As CSectionModel
    FocusedSelectGeometry "SavedWorkbook"
    SetSystemSetting "Load.ReferenceOffsetX", "0": SetSystemSetting "Load.ReferenceOffsetY", "0"
    SetSystemSetting "Calculation.ZeroMomentPerDepth", "0"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    Set loadRange = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ReDim values(1 To loadRange.Rows.Count - 1, 1 To loadRange.Columns.Count)
    For Each item In FocusedSelectedLoads(section, "Generated")
        If CStr(item(0)) = "My-1" Then
            values(1, 1) = "My-1": values(1, 2) = 0#: values(1, 3) = 0#
            values(1, 4) = CDbl(item(3)) / units.InputMomentMyToInternal(1#)
            values(1, 5) = "PR3": values(1, 6) = "Auto"
        End If
    Next item
    loadRange.Offset(1, 0).Resize(loadRange.Rows.Count - 1, loadRange.Columns.Count).Value2 = values
    SetProfileValue "Calculation.Strength.DirectState", "PR3", "Yes"
    SetProfileValue "Calculation.Strength.Capacity", "PR3", "Yes"
    SetProfileValue "Calculation.Crack.Width", "PR3", "Yes"
    SetProfileValue "Calculation.Stability.Enabled", "PR3", "No"
    output = RunSectionCalculationForWorkbook(ThisWorkbook, False) & vbCrLf
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set state = mFocusedBatch.ResultAt(1).StateRepository.FindState(sstCapacityState)
    Set concrete = provider.ConcreteMaterialForEquilibriumFromSpec(state.MaterialSpec)
    Set steel = provider.SteelMaterialForEquilibriumFromSpec(state.MaterialSpec)
    SnapshotDiagnosticPoints output, "actualC", concrete
    SnapshotDiagnosticPoints output, "savedC", state.ConcreteDiagramSnapshot
    SnapshotDiagnosticPoints output, "actualS", steel
    SnapshotDiagnosticPoints output, "savedS", state.SteelDiagramSnapshot
    Set probe = New CSectionSolver
    probe.EvaluateStrainPlane mFocusedSection, concrete, steel, state.Epsilon0, state.KappaX, state.KappaY
    output = output & "DIAGNOSTIC_ACTUAL_FORCE|N=" & SnapshotDiagnosticNumber(probe.Nint) & _
        "|Mx=" & SnapshotDiagnosticNumber(probe.Mxint) & "|My=" & SnapshotDiagnosticNumber(probe.Myint) & vbCrLf
    Set probe = New CSectionSolver
    probe.EvaluateStrainPlane mFocusedSection, state.ConcreteDiagramSnapshot, state.SteelDiagramSnapshot, state.Epsilon0, state.KappaX, state.KappaY
    output = output & "DIAGNOSTIC_SAVED_FORCE|N=" & SnapshotDiagnosticNumber(probe.Nint) & _
        "|Mx=" & SnapshotDiagnosticNumber(probe.Mxint) & "|My=" & SnapshotDiagnosticNumber(probe.Myint) & vbCrLf
    mFocusedSection.ConcreteMeshCentroid cx, cy, area
    centerStrain = state.Epsilon0 + state.KappaX * cy + state.KappaY * cx
    For i = 1 To mFocusedSection.ConcreteCount
        strain = centerStrain + state.KappaX * (mFocusedSection.ConcreteY(i) - cy) + state.KappaY * (mFocusedSection.ConcreteX(i) - cx)
        concrete.EvaluateAtStrain strain, actual, tangent
        state.ConcreteDiagramSnapshot.EvaluateAtStrain strain, saved, savedTangent
        If actual <> saved Then output = output & "DIAGNOSTIC_STRESS|kind=C|id=" & CStr(i) & _
            "|strain=" & SnapshotDiagnosticNumber(strain) & "|actual=" & SnapshotDiagnosticNumber(actual) & "|saved=" & SnapshotDiagnosticNumber(saved) & vbCrLf
    Next i
    For i = 1 To mFocusedSection.RebarCount
        strain = centerStrain + state.KappaX * (mFocusedSection.RebarY(i) - cy) + state.KappaY * (mFocusedSection.RebarX(i) - cx)
        steel.EvaluateAtStrain strain, actual, tangent
        state.SteelDiagramSnapshot.EvaluateAtStrain strain, saved, savedTangent
        If actual <> saved Then output = output & "DIAGNOSTIC_STRESS|kind=S|id=" & CStr(i) & _
            "|strain=" & SnapshotDiagnosticNumber(strain) & "|actual=" & SnapshotDiagnosticNumber(actual) & "|saved=" & SnapshotDiagnosticNumber(saved) & vbCrLf
        concrete.EvaluateAtStrain strain, actual, tangent
        state.ConcreteDiagramSnapshot.EvaluateAtStrain strain, saved, savedTangent
        If actual <> saved Then output = output & "DIAGNOSTIC_STRESS|kind=DISPLACED_C|id=" & CStr(i) & _
            "|strain=" & SnapshotDiagnosticNumber(strain) & "|actual=" & SnapshotDiagnosticNumber(actual) & "|saved=" & SnapshotDiagnosticNumber(saved) & vbCrLf
    Next i
    RunFocusedSnapshotDiagnostic = output
    Exit Function
Failed:
    RunFocusedSnapshotDiagnostic = "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

Private Sub SnapshotDiagnosticPoints(ByRef output As String, ByVal label As String, ByVal material As CMaterialDiagram)
    Dim i As Long
    For i = 1 To material.PointCount
        output = output & "DIAGNOSTIC_POINT|label=" & label & "|strain=" & SnapshotDiagnosticNumber(material.PointStrain(i)) & _
            "|stress=" & SnapshotDiagnosticNumber(material.PointStress(i)) & vbCrLf
    Next i
End Sub

Private Function SnapshotDiagnosticNumber(ByVal value As Double) As String
    SnapshotDiagnosticNumber = Replace$(Format$(value, "0.0000000000000000E+00"), ",", ".")
End Function
