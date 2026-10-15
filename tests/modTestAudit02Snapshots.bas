Attribute VB_Name = "modTestAudit02Snapshots"
Option Explicit

' ДЛЯ ТЕСТОВ: сравнение построения и направленные расчёты всех групп сечений.
#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If
Private Type TDoubleValue
    Value As Double
End Type
Private Type TDoubleWords
    Low As Long
    High As Long
End Type

' ДЛЯ ТЕСТОВ: один прогрев и пять замеров полного построения модели каждого варианта геометрии.
Public Function DumpAudit02BuildPerformance() As String
    Dim config As Object, saved As Variant, fixture As Variant, i As Long, result As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, model As CSectionModel
    Dim started As Currency, finished As Currency, frequency As Currency
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    If QueryPerformanceFrequency(frequency) = 0 Then Err.Raise vbObjectError + 5490, , "Недоступен точный таймер."
    For Each fixture In Array("Circle", "Rectangle", "LSection", "TwoRectangles", "RoundedSimple", _
            "RoundedTapered", "RoundedMixed", "RoundedMixedLeft", "HollowCentered", "HollowOffset")
        config.Formula = saved: modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
        For i = 1 To 5
            QueryPerformanceCounter started
            Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
            QueryPerformanceCounter finished
            result = result & CStr(fixture) & "|" & CStr(i) & "|" & CStr((finished - started) / frequency) & _
                "|" & CStr(model.ConcreteCount) & "|" & CStr(model.RebarCount) & vbCrLf
        Next i
    Next fixture
    GoTo Finished
Failed:
    result = result & "ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    DumpAudit02BuildPerformance = result
End Function

' ДЛЯ ТЕСТОВ: два сочетания четырёх групп сечений с НДС, несущей способностью,
' устойчивостью и трещинами. Сохраняет все биты чисел Double, статусы
' и пояснения отдельно. Настройку продолжения диаграмм не меняет.
Public Function DumpAudit02Calculations(Optional ByVal crackCode As String = "SP35") As String
    Dim config As Object, saved As Variant, fixture As Variant, result As String
    Dim profiles As Object, row As Long, key As String, settings As CSystemSettingsReader, units As CUnitSystem
    Dim model As CSectionModel, properties As CSectionPropertiesCalculator, provider As CMaterialModelProvider
    Dim catalog As CCalculationProfileCatalog, batch As CBatchSectionCalculator, item As CCombinationResult
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double, i As Long, meta As CResultMeta
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For Each fixture In Array("Circle", "LSection", "RoundedTapered", "HollowOffset")
        config.Formula = saved
        modTestShapeArchitecture.ConfigureShapeFixture "Circle"
        modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
        Set profiles = ThisWorkbook.Names("rngCalculationProfiles").RefersToRange
        For row = 2 To profiles.Rows.Count
            key = CStr(profiles.Cells(row, 2).Value2)
            Select Case key
                Case "Calculation.Strength.DirectState", "Calculation.Strength.Capacity", "Calculation.Stability.Enabled", _
                     "Calculation.Crack.Width", "Calculation.Crack.Longitudinal"
                    profiles.Cells(row, 3).Value2 = "Yes"
            End Select
        Next row
        SetSetting "SLS.Crack.Code", crackCode
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
        Set properties = New CSectionPropertiesCalculator: properties.PrepareSnapshot model
        properties.CalculateProjection model, 1#, 0#, False, minX, maxX
        properties.CalculateProjection model, 0#, 1#, False, minY, maxY
        Set provider = New CMaterialModelProvider: provider.Initialize settings, units
        Set catalog = New CCalculationProfileCatalog: catalog.LoadFromWorkbook ThisWorkbook
        Set batch = New CBatchSectionCalculator: batch.Initialize model, provider
        Set batch.ProfileCatalog = catalog: batch.ApplySettings settings, units
        batch.ApplyLoadReference properties.CentroidX, properties.CentroidY, 0#, 0#
        batch.AddCombination "COMP", -2# * properties.Area, 0#, 0#, "PR1", "Audit02 control", "Auto"
        batch.AddCombination "BEND", -0.2 * properties.Area, properties.Area * (maxY - minY) * 1.6, _
            -properties.Area * (maxX - minX) * 0.6, "PR1", "Audit02 control", "Auto"
        batch.Execute
        For i = 1 To batch.Count
            Set item = batch.ResultAt(i)
            result = result & CStr(fixture) & "|" & CStr(i) & "|STATUS|" & item.Status & "|" & item.NormalCrackStatus & _
                "|" & CStr(item.DirectStateMeta.InternalStatus) & "|" & CStr(item.CapacityMeta.InternalStatus) & _
                "|" & CStr(item.StabilityMeta.InternalStatus) & "|" & CStr(item.CrackFormationMeta.InternalStatus) & _
                "|" & CStr(item.CrackWidthMeta.InternalStatus) & "|" & CStr(item.LongitudinalCrackMeta.InternalStatus) & vbCrLf
            result = result & CStr(fixture) & "|" & CStr(i) & "|VALUES|" & Bits(item.StrengthResult.Capacity.LambdaCapacity) & _
                "|" & Bits(item.CrackResult.Formation.LambdaCrc) & "|" & Bits(item.CrackResult.Width.CrackWidth) & _
                "|" & Bits(item.CrackResult.Width.Abt) & "|" & Bits(item.StabilityResult.CriticalForce) & _
                "|" & Bits(item.StabilityResult.Eta1) & "|" & Bits(item.StabilityResult.Eta2) & _
                "|" & Bits(item.StabilityResult.DesignN) & "|" & Bits(item.StabilityResult.DesignMx) & "|" & Bits(item.StabilityResult.DesignMy) & vbCrLf
            Set meta = item.OverallMeta
            result = result & CStr(fixture) & "|" & CStr(i) & "|COMMENT|" & Replace$(meta.ResultComment, vbCrLf, " / ") & vbCrLf
            result = result & CStr(fixture) & "|" & CStr(i) & "|STABILITY|" & Replace$(item.StabilityMeta.ResultComment, vbCrLf, " / ") & vbCrLf
        Next i
    Next fixture
    GoTo Finished
Failed:
    result = result & "ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    DumpAudit02Calculations = result
End Function

' Изменяет только адресный параметр в проверочной копии, без фиксированной координаты Config.
Private Sub SetSetting(ByVal key As String, ByVal value As String)
    Dim table As Object, row As Long
    Set table = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = key Then table.Cells(row, 2).Value2 = value: Exit Sub
    Next row
    Err.Raise vbObjectError + 5491, , "Нет параметра " & key
End Sub
Private Function Bits(ByVal value As Double) As String
    Dim number As TDoubleValue, words As TDoubleWords
    number.Value = value: LSet words = number
    Bits = Right$("00000000" & Hex$(words.High), 8) & Right$("00000000" & Hex$(words.Low), 8)
End Function
