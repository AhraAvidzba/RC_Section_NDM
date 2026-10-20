Attribute VB_Name = "modTestZeroContourParity"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: побитовое сравнение геометрии при a2 = a3 = 0
' ==========================================================================
' Этот модуль использует только прежний API и выполняется в исходной и новой
' книге. Каждый Double выводится восемью исходными байтами, без округления.

Private Type TDoubleValue
    Value As Double ' Исходное значение координаты или геометрического свойства.
End Type
Private Type TDoubleBytes
    Byte1 As Byte ' Первый байт исходного Double в памяти.
    Byte2 As Byte ' Второй байт исходного Double в памяти.
    Byte3 As Byte ' Третий байт исходного Double в памяти.
    Byte4 As Byte ' Четвёртый байт исходного Double в памяти.
    Byte5 As Byte ' Пятый байт исходного Double в памяти.
    Byte6 As Byte ' Шестой байт исходного Double в памяти.
    Byte7 As Byte ' Седьмой байт исходного Double в памяти.
    Byte8 As Byte ' Восьмой байт исходного Double в памяти.
End Type

' ДЛЯ ТЕСТОВ: пять форм, оба расположения связанных рядов и обе привязки.
Public Function DumpZeroContourParity() As String
    Dim config As Object, saved As Variant, fixture As Variant, tableName As Variant, key As Variant, area As Object
    Dim keyed As CSectionInputReader, settings As CSystemSettingsReader, units As CUnitSystem
    Dim builder As CSectionBuilder, model As CSectionModel, scenario As Long, result As String, count As Long
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = FixtureSaveInputs(config)
    For Each fixture In Array("Circle", "Rectangle", "LSection", "RoundedMixed", "HollowOffset")
        For scenario = 0 To 3
            modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
            For Each tableName In Array("rngCircleGeometry", "rngRectanglesGeometry", "rngPointedSectionGeometry", "rngHollowRectangleGeometry")
                Set area = ThisWorkbook.Names(CStr(tableName)).RefersToRange: Set keyed = FixtureInput(area.Value2)
                For Each key In keyed.RowKeys("Extra")
                    FixtureInputCell(area, "Extra", CStr(key), "a2").Value2 = 0#
                    FixtureInputCell(area, "Extra", CStr(key), "a3").Value2 = 0#
                    FixtureInputCell(area, "Extra", CStr(key), "pos2").Value2 = IIf(scenario >= 2, "SideBySide", "Stacked")
                    FixtureInputCell(area, "Extra", CStr(key), "pos3").Value2 = IIf(scenario >= 2, "SideBySide", "Stacked")
                    FixtureInputCell(area, "Extra", CStr(key), "bind2").Value2 = IIf(scenario Mod 2 = 1, "EverySecondBar", "EachBar")
                    FixtureInputCell(area, "Extra", CStr(key), "bind3").Value2 = IIf(scenario Mod 2 = 1, "EverySecondBar", "EachBar")
                Next key
            Next tableName
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
            result = result & ModelBits(CStr(fixture) & "." & CStr(scenario), model)
            count = count + 1
        Next scenario
    Next fixture
    result = result & "ZERO_PARITY_CASES: " & CStr(count) & vbCrLf
    GoTo Finished
Failed:
    result = result & "FAIL: zero.parity; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then FixtureRestoreInputs config, saved
    DumpZeroContourParity = result
End Function

' Сохраняет имена и все геометрические числа бетона и арматуры в исходном порядке.
Private Function ModelBits(ByVal name As String, ByVal model As CSectionModel) As String
    Dim lines() As String, i As Long, at As Long
    ReDim lines(0 To model.ConcreteCount + model.RebarCount)
    lines(0) = "ZERO_CASE: " & name & "; concrete=" & CStr(model.ConcreteCount) & "; rebar=" & CStr(model.RebarCount)
    For i = 1 To model.ConcreteCount
        at = at + 1
        lines(at) = "C|" & model.ConcreteID(i) & "|" & Bits(model.ConcreteX(i)) & "|" & Bits(model.ConcreteY(i)) & _
            "|" & Bits(model.ConcreteArea(i)) & "|" & Bits(model.ConcreteLocalIx(i)) & "|" & Bits(model.ConcreteLocalIy(i)) & _
            "|" & Bits(model.ConcreteLocalIxy(i)) & "|" & Bits(model.ConcreteWidth(i)) & "|" & Bits(model.ConcreteHeight(i)) & "|" & Bits(model.ConcreteRotation(i))
    Next i
    For i = 1 To model.RebarCount
        at = at + 1
        lines(at) = "R|" & model.RebarID(i) & "|" & model.RebarSourceName(i) & "|" & Bits(model.RebarX(i)) & _
            "|" & Bits(model.RebarY(i)) & "|" & Bits(model.RebarArea(i)) & "|" & Bits(model.RebarDiameter(i))
    Next i
    ModelBits = Join(lines, vbCrLf) & vbCrLf
End Function

' Преобразует Double в 16 шестнадцатеричных знаков без изменения его битов.
Private Function Bits(ByVal value As Double) As String
    Dim number As TDoubleValue, bytes As TDoubleBytes
    number.Value = value: LSet bytes = number
    Bits = Right$("0" & Hex$(bytes.Byte1), 2) & Right$("0" & Hex$(bytes.Byte2), 2) & _
        Right$("0" & Hex$(bytes.Byte3), 2) & Right$("0" & Hex$(bytes.Byte4), 2) & _
        Right$("0" & Hex$(bytes.Byte5), 2) & Right$("0" & Hex$(bytes.Byte6), 2) & _
        Right$("0" & Hex$(bytes.Byte7), 2) & Right$("0" & Hex$(bytes.Byte8), 2)
End Function
