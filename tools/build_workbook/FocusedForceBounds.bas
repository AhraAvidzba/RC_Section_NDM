' ДЛЯ ТЕСТОВ: фактическая геометрия и диаграммы PR3 для независимых границ усилий.
Public Function DumpFocusedForceBounds(Optional ByVal shapeName As String = "SavedWorkbook") As String
    On Error GoTo Failed
    Dim settings As CSystemSettingsReader, units As CUnitSystem, section As CSectionModel
    Dim profiles As CCalculationProfileCatalog, provider As CMaterialModelProvider, spec As CMaterialModelSpec
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, mode As Variant, label As Variant
    Dim output As String, x As Double, y As Double, i As Long, oldSystem As Variant
    oldSystem = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set section = Audit03LoadMatrixSection(shapeName)
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    CalculateConcreteSectionCentroid section, x, y
    output = "CENTER|x=" & FocusedBoundNumber(x) & "|y=" & FocusedBoundNumber(y) & vbCrLf
    For i = 1 To section.ConcreteCount
        output = output & "FIBER|kind=C|x=" & FocusedBoundNumber(section.ConcreteX(i) - x) & _
            "|y=" & FocusedBoundNumber(section.ConcreteY(i) - y) & "|area=" & FocusedBoundNumber(section.ConcreteArea(i)) & vbCrLf
    Next i
    For i = 1 To section.RebarCount
        output = output & "FIBER|kind=S|x=" & FocusedBoundNumber(section.RebarX(i) - x) & _
            "|y=" & FocusedBoundNumber(section.RebarY(i) - y) & "|area=" & FocusedBoundNumber(section.RebarArea(i)) & vbCrLf
    Next i
    For Each mode In Array("No", "Yes")
        SetSystemSetting "General.DiagramExtension", CStr(mode)
        settings.LoadFromWorkbook ThisWorkbook
        Set provider = New CMaterialModelProvider: provider.Initialize settings, units
        For Each label In Array("PR1", "PR2", "PR3")
            Select Case CStr(label)
                Case "PR1": Set spec = profiles.ProfileById("PR3").StrengthSpec
                Case "PR2": Set spec = profiles.ProfileById("PR3").CrackedStateSpec
                Case "PR3": Set spec = profiles.ProfileById("PR3").CrackInitiationSpec
            End Select
            Set concrete = provider.ConcreteMaterialForEquilibriumFromSpec(spec)
            Set steel = provider.SteelMaterialForEquilibriumFromSpec(spec)
            FocusedBoundPoints output, CStr(mode), CStr(label), "C", concrete
            FocusedBoundPoints output, CStr(mode), CStr(label), "S", steel
        Next label
    Next mode
    GoTo Restore
Failed:
    output = "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
Restore:
    If IsArray(oldSystem) Then ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula = oldSystem
    DumpFocusedForceBounds = output
End Function

Private Sub FocusedBoundPoints(ByRef output As String, ByVal mode As String, ByVal profile As String, _
        ByVal kind As String, ByVal material As CMaterialDiagram)
    Dim i As Long
    For i = 1 To material.PointCount
        output = output & "MATERIAL|mode=" & mode & "|profile=" & profile & "|kind=" & kind & _
            "|strain=" & FocusedBoundNumber(material.PointStrain(i)) & "|stress=" & FocusedBoundNumber(material.PointStress(i)) & vbCrLf
    Next i
End Sub

Private Function FocusedBoundNumber(ByVal value As Double) As String
    FocusedBoundNumber = Replace$(Format$(value, "0.0000000000000000E+00"), ",", ".")
End Function
