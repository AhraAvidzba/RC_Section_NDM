Attribute VB_Name = "modTestSectionInputReader"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: Проверки общего ридера: перестановки, типы значений, единицы и адреса ошибок.
' ==========================================================================
Private mPassed As Long, mFailed As Long, mReport As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт два независимых блока; ошибочное числовое поле остаётся невостребованным.
Private Function TestInputValues() As Variant
    Dim data(1 To 13, 1 To 6) As Variant
    data(1, 1) = "Dimensions": data(1, 2) = "Размеры"
    data(3, 1) = "key": data(3, 2) = "name": data(3, 3) = "val"
    data(3, 4) = "unit": data(3, 5) = "note"
    data(4, 1) = "B": data(4, 2) = "Ширина": data(4, 3) = 60#
    data(5, 1) = "Zero": data(5, 3) = 0#
    data(6, 1) = "Unused": data(6, 3) = CVErr(2015)
    data(8, 1) = "Bars": data(8, 2) = "Армирование"
    data(10, 1) = "key": data(10, 2) = "name": data(10, 3) = "d"
    data(10, 4) = "n": data(10, 5) = "mode": data(10, 6) = "as"
    data(11, 1) = "B": data(11, 3) = 16#: data(11, 4) = 4
    data(11, 5) = "stacked": data(11, 6) = 40#
    data(12, 1) = "Off": data(12, 4) = 0
    data(13, 1) = "Units": data(13, 3) = "мм"
    TestInputValues = data
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит снимок с адресами, которые зависят от фактической позиции, а не от ключа.
Private Function TestMakeReader(ByVal data As Variant, Optional ByVal units As CUnitSystem = Nothing) As CSectionInputReader
    Dim locations(1 To 13, 1 To 6) As String, row As Long, column As Long
    For row = 1 To 13
        For column = 1 To 6
            locations(row, column) = "Config!" & Chr$(71 + column) & CStr(80 + row)
        Next column
    Next row
    Dim table As CSectionInputTable: Set table = New CSectionInputTable
    table.Initialize "testTable", data, locations
    Set TestMakeReader = New CSectionInputReader: TestMakeReader.Initialize table, units
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет ошибку структуры либо выбранного числового поля и точный адрес её причины.
Private Sub TestReject(ByVal title As String, ByVal data As Variant, ByVal expectedAddress As String, Optional ByVal operation As String = "structure")
    Dim reader As CSectionInputReader, ignored As Double, description As String, keys As Variant
    On Error Resume Next
    Set reader = TestMakeReader(data)
    If Err.Number = 0 Then
        Select Case operation
            Case "structure": keys = reader.RowKeys("Dimensions")
            Case "count": ignored = reader.Count("Bars", "B", "n")
            Case "diameter": ignored = reader.OptionalDiameter("Bars", "B", "d")
            Case "value": ignored = reader.Number("Dimensions", "Unused")
            Case "missing": ignored = reader.Number("Dimensions", "Missing")
            Case "column": ignored = reader.Number("Dimensions", "B", "Missing")
            Case "block": ignored = reader.Number("Missing", "B")
            Case "barKeys": keys = reader.RowKeys("Bars")
        End Select
    End If
    description = Err.Description
    TestCheck title, Err.Number <> 0 And InStr(1, description, expectedAddress, vbTextCompare) > 0
    Err.Clear: On Error GoTo 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает результат отдельного условия, сохраняя все причины неудач.
Private Sub TestCheck(ByVal title As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & title & vbCrLf
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Проверяет независимость чтения от подписей, порядка строк, столбцов и блоков.
Public Function RunSectionInputReaderTests() As String
    Dim data As Variant, copy As Variant, reader As CSectionInputReader, scenario As Long
    Dim row As Long, column As Long, value As Variant, units As CUnitSystem, key As Variant, scoped As CSectionInputReader, scopedTable As CSectionInputTable
    Dim custom(1 To 2, 1 To 2) As Variant, customTable As CSectionInputTable, ignored As Variant, code As Long, description As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Set units = New CUnitSystem: units.InitializeDefaults
    data = TestInputValues: data(11, 1) = "TypeA.B": data(12, 1) = "TypeB.Off"
    Set scopedTable = New CSectionInputTable: scopedTable.Initialize "sharedInput", data
    Set scoped = New CSectionInputReader: scoped.Initialize scopedTable, units, "typea."
    TestCheck "prefix.count", scoped.Count("Bars", "B", "n") = 4
    key = scoped.RowKeys("Bars")
    TestCheck "prefix.keys", UBound(key) = 0 And CStr(key(0)) = "B"
    TestCheck "prefix.column", scoped.ColumnIndex("Bars", "d") = 3
    Set scoped = New CSectionInputReader: scoped.Initialize scopedTable, units, "Missing."
    key = scoped.RowKeys("Bars")
    TestCheck "prefix.empty", UBound(key) < LBound(key)
    For scenario = 0 To 7 ' Три бита независимо переставляют столбцы, строки данных и целые блоки: проверяются все восемь сочетаний.
        data = TestInputValues
        If (scenario And 1) <> 0 Then
            For row = 1 To 13
                value = data(row, 3): data(row, 3) = data(row, 6): data(row, 6) = value
            Next row
        End If
        If (scenario And 2) <> 0 Then
            For column = 1 To 6
                value = data(4, column): data(4, column) = data(5, column): data(5, column) = value
                value = data(11, column): data(11, column) = data(12, column): data(12, column) = value
            Next column
        End If
        If (scenario And 4) <> 0 Then
            For row = 1 To 6
                For column = 1 To 6
                    value = data(row, column): data(row, column) = data(row + 7, column): data(row + 7, column) = value
                Next column
            Next row
        End If
        Set reader = TestMakeReader(data, units)
        TestCheck "number." & CStr(scenario), reader.Number("dimensions", "b") = 60#
        TestCheck "length." & CStr(scenario), reader.PositiveLength("Dimensions", "B") = 60#
        TestCheck "count." & CStr(scenario), reader.Count("Bars", "B", "n") = 4
        TestCheck "diameter." & CStr(scenario), reader.OptionalDiameter("Bars", "B", "d") = 16#
        TestCheck "choice." & CStr(scenario), reader.Choice("Bars", "B", Array("Stacked", "SideBySide"), "mode") = "Stacked"
        TestCheck "empty.diameter." & CStr(scenario), reader.OptionalDiameter("Bars", "Off", "d") = 0#
        TestCheck "empty.offset." & CStr(scenario), reader.OptionalRowOffset("Bars", "Off", "as") = 0#
        key = reader.RowKeys("Bars"): TestCheck "units.excluded." & CStr(scenario), UBound(key) = 1
        copy = data: data(1, 2) = "Изменённый заголовок"
        TestCheck "snapshot." & CStr(scenario), reader.Number("Dimensions", "B") = 60#
    Next scenario
    data = TestInputValues: data(8, 1) = "Dimensions": TestReject "duplicate.block", data, "Config!H88"
    data = TestInputValues: data(5, 1) = "B": TestReject "duplicate.row", data, "Config!H85"
    data = TestInputValues: data(3, 4) = "val": TestReject "duplicate.column", data, "Config!K83"
    data = TestInputValues: TestReject "unused.error.lazy", data, "Config!J86", "value"
    data = TestInputValues: data(11, 1) = CVErr(2015): Set reader = TestMakeReader(data)
    TestCheck "unused.block.key.error.lazy", reader.Number("Dimensions", "B") = 60#
    TestReject "requested.block.key.error.address", data, "Config!H91", "barKeys"
    TestReject "missing.row", data, "Config!H81", "missing"
    TestReject "missing.column", data, "Config!H83", "column"
    TestReject "missing.block", data, "Config!H81", "block"
    For Each value In Array(-1, 2.5, True, "text")
        data = TestInputValues: data(11, 4) = value: TestReject "invalid.count", data, "Config!K91", "count"
    Next value
    For Each value In Array(-1, "-", True, "text")
        data = TestInputValues: data(11, 3) = value: TestReject "invalid.diameter", data, "Config!J91", "diameter"
    Next value
    Set reader = TestMakeReader(TestInputValues, units)
    units.CopyState "cm", "m", "mm2", "m2", "N", "N", "N*mm", "N*mm", "MPa", "MPa", "1/mm", "1/mm", "Tension", "+Y tension", "+X tension"
    TestCheck "units.snapshot.unchanged", reader.PositiveLength("Dimensions", "B") = 60#
    Set reader = TestMakeReader(TestInputValues, units)
    TestCheck "units.cm.to.mm", reader.PositiveLength("Dimensions", "B") = 600#
    TestCheck "raw.value.no.conversion", reader.CellValue("Dimensions", "B") = 60#
    TestCheck "text.original.case", reader.Text("Bars", "B", "mode") = "stacked"
    TestCheck "signed.length.cm.to.mm", reader.Length("Dimensions", "B") = 600#
    TestCheck "nonnegative.zero", reader.NonNegativeLength("Dimensions", "Zero") = 0#
    TestCheck "source.related.fields", reader.SourceReferences("Bars", "B", Array("d", "n")) = "Config!J91; Config!K91"
    reader.Require "Dimensions", "B", "val", True, "Размер допустим."
    TestCheck "require.valid", True
    On Error Resume Next
    reader.Require "Dimensions", "B", "val", False, "Ширина недостаточна."
    description = Err.Description: code = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "require.actual.address", code <> 0 And InStr(description, "Config!J84") > 0 And InStr(description, "Ширина недостаточна.") > 0
    custom(1, 1) = "Размер": custom(1, 2) = "Значение": custom(2, 1) = "B": custom(2, 2) = 25#
    Set customTable = New CSectionInputTable: customTable.Initialize "Custom", custom
    Set reader = New CSectionInputReader: reader.Initialize customTable, units
    TestCheck "custom.positional.text", reader.ValueAt(1, 2) = "Значение"
    TestCheck "custom.positional.number.raw", reader.ValueAt(2, 2) = 25#
    custom(2, 2) = 99#: TestCheck "custom.positional.snapshot", reader.ValueAt(2, 2) = 25#
    On Error Resume Next
    ignored = reader.ValueAt(3, 1): code = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "custom.positional.bounds", code <> 0
    custom(2, 2) = CVErr(2015)
    Set customTable = New CSectionInputTable: customTable.Initialize "Custom", custom
    Set reader = New CSectionInputReader: reader.Initialize customTable
    On Error Resume Next
    ignored = reader.ValueAt(2, 2): code = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo Failed
    TestCheck "custom.positional.formula.error", code <> 0 And InStr(description, "Custom[2,2]") > 0
    GoTo Finished
Failed:
    mFailed = mFailed + 1: mReport = mReport & "RUNTIME ERROR: " & Err.Description & vbCrLf
Finished:
    RunSectionInputReaderTests = mReport & "TOTAL_SECTION_INPUT_READER: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function
