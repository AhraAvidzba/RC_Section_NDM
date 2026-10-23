' ==========================================================================
' ДЛЯ ТЕСТОВ — вспомогательный фрагмент частной приёмки; в выпуск не включается.
' ==========================================================================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сериализует фактические волокна, диаграммы PR3 и допуски для независимой оценки усилий без решения НДС.
Public Function TestDumpFocusedForceBounds(Optional ByVal shapeName As String = "SavedWorkbook") As String
    Dim oldRebars As Variant
    oldRebars = ThisWorkbook.Worksheets("Сечения").UsedRange.Formula
    On Error GoTo Failed
    Dim settings As CSystemSettingsReader, units As CUnitSystem, section As CSectionModel
    Dim profiles As CCalculationProfileCatalog, provider As CMaterialModelProvider, spec As CMaterialModelSpec
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, mode As Variant, label As Variant
    Dim output As String, x As Double, y As Double, i As Long, oldSystem As Variant
    oldSystem = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula
    TestFocusedSelectGeometry shapeName
    TestSetSystemSetting "SLS.Crack.Code", "SP35"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    CalculateConcreteSectionCentroid section, x, y
    output = "CENTER|x=" & TestFocusedBoundNumber(x) & "|y=" & TestFocusedBoundNumber(y) & vbCrLf
    output = output & "ACCEPTANCE_TOLERANCE|N=" & TestFocusedBoundNumber(Abs(units.InputForceToInternal(settings.GetDouble("Solver.ToleranceN")))) & _
        "|Mx=" & TestFocusedBoundNumber(Abs(units.InputMomentMxToInternal(settings.GetDouble("Solver.ToleranceMx")))) & _
        "|My=" & TestFocusedBoundNumber(Abs(units.InputMomentMyToInternal(settings.GetDouble("Solver.ToleranceMy")))) & vbCrLf
    For i = 1 To section.ConcreteCount
        output = output & "FIBER|kind=C|x=" & TestFocusedBoundNumber(section.ConcreteX(i) - x) & _
            "|y=" & TestFocusedBoundNumber(section.ConcreteY(i) - y) & "|area=" & TestFocusedBoundNumber(section.ConcreteArea(i)) & vbCrLf
    Next i
    For i = 1 To section.RebarCount
        output = output & "FIBER|kind=S|x=" & TestFocusedBoundNumber(section.RebarX(i) - x) & _
            "|y=" & TestFocusedBoundNumber(section.RebarY(i) - y) & "|area=" & TestFocusedBoundNumber(section.RebarArea(i)) & vbCrLf
    Next i
    For Each mode In Array("No", "Yes")
        TestSetSystemSetting "General.DiagramExtension", CStr(mode)
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
            TestFocusedBoundPoints output, CStr(mode), CStr(label), "C", concrete
            TestFocusedBoundPoints output, CStr(mode), CStr(label), "S", steel
        Next label
    Next mode
    GoTo Restore
Failed:
    output = "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
Restore:
    If IsArray(oldRebars) Then ThisWorkbook.Worksheets("Сечения").UsedRange.Formula = oldRebars
    If IsArray(oldSystem) Then ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula = oldSystem
    TestDumpFocusedForceBounds = output
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает физические пределы, признаки extension и все точки фактической диаграммы для независимого интеграла.
Private Sub TestFocusedBoundPoints(ByRef output As String, ByVal mode As String, ByVal profile As String, _
        ByVal kind As String, ByVal material As CMaterialDiagram)
    Dim i As Long
    output = output & "MATERIAL_FLAGS|mode=" & mode & "|profile=" & profile & "|kind=" & kind & _
        "|compressionExtension=" & CStr(material.HasCompressionExtension) & "|tensionExtension=" & CStr(material.HasTensionExtension) & _
        "|compressionLimit=" & TestFocusedBoundNumber(material.PhysicalCompressionStrain) & _
        "|tensionLimit=" & TestFocusedBoundNumber(material.PhysicalTensionStrain) & vbCrLf
    For i = 1 To material.PointCount
        output = output & "MATERIAL|mode=" & mode & "|profile=" & profile & "|kind=" & kind & _
            "|strain=" & TestFocusedBoundNumber(material.PointStrain(i)) & "|stress=" & TestFocusedBoundNumber(material.PointStress(i)) & vbCrLf
    Next i
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выбирает F64 либо инвариантную научную запись чисел модели, не округляя сами расчётные данные.
Private Function TestFocusedBoundNumber(ByVal value As Double) As String
    If mFocusedBinaryDump Then
        TestFocusedBoundNumber = TestFocusedExactNumber(value)
    Else
        TestFocusedBoundNumber = Replace$(Format$(value, "0.0000000000000000E+00"), ",", ".")
    End If
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Включает F64-формат при выгрузке фактической модели для независимой проверки больших усилий.
Public Function TestDumpFocusedForceBoundsExact(Optional ByVal shapeName As String = "SavedWorkbook") As String
    mFocusedBinaryDump = True
    TestDumpFocusedForceBoundsExact = TestDumpFocusedForceBounds(shapeName)
    mFocusedBinaryDump = False
End Function
