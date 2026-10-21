Attribute VB_Name = "WorkflowEquivalenceProbe"
Option Explicit


' Фиксирует этапы ограниченной проверки до возврата большого массива результатов.
Private Sub Progress(ByVal value As String)
    Dim file As Integer: file = FreeFile
    Open ThisWorkbook.Path & "\Audit03_Search_Progress.txt" For Append As #file
    Print #file, Format$(Now, "hh:nn:ss") & "|" & value
    Close #file
End Sub

' ДЛЯ ТЕСТОВ: одинаковые исходные данные для прежней и новой последовательности.
' Модуль добавляется только в изолированные копии; рабочие алгоритмы не меняет.
Private Sub Setting(ByVal key As String, ByVal value As Variant)
    Dim table As Object, data As Variant, row As Long
    Set table = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    data = table.Value2
    For row = 1 To UBound(data, 1)
        If CStr(data(row, 1)) = key Then table.Cells(row, 2).Value2 = value: Exit Sub
    Next row
    Err.Raise 5, , "Не найдена настройка " & key
End Sub

' Приводит строки и числа к однострочной записи без потери полей сообщений.
Private Function Text(ByVal value As Variant) As String
    If IsError(value) Then Err.Raise 5, , "Ошибка Excel в сохранённом результате."
    Text = Replace(Replace(Replace(CStr(value), "\", "\\"), "|", "\p"), vbCrLf, "\n")
    Text = Replace(Replace(Text, vbLf, "\n"), vbCr, "\n")
End Function


' Читает один блок в копии старой книги, где общий reader был закрытым.
Private Function ReadTable(ByVal name As String) As Variant
    Dim anchor As Object, headers As Variant, width As Long, rows As Long
    Set anchor = ThisWorkbook.Names(name).RefersToRange
    headers = anchor.Resize(1, 32).Value2
    For width = 1 To 32
        If Len(CStr(headers(1, width))) = 0 Then Exit For
    Next width
    width = width - 1
    rows = anchor.Worksheet.Cells(anchor.Worksheet.Rows.Count, anchor.Column).End(-4162).Row - anchor.Row + 1
    If rows < 1 Or width < 1 Then Err.Raise 5, , "Пустой блок " & name
    ReadTable = anchor.Resize(rows, width).Value2
End Function

' Выгружает массив результата; служебная дата запуска не сравнивается.
Private Function Dump(ByVal fixture As String, ByVal name As String, ByVal data As Variant, _
        Optional ByVal summary As Boolean = False) As String
    Dim row As Long, column As Long, firstColumn As Long, lines() As String, count As Long
    ReDim lines(1 To UBound(data, 1) * UBound(data, 2))
    firstColumn = 2: If summary Then firstColumn = 1
    For row = 1 To UBound(data, 1)
        For column = firstColumn To UBound(data, 2)
            If Not summary Or row <> 4 Then
                count = count + 1
                lines(count) = "EQ|" & fixture & "|" & name & "|" & row & "|" & column & "|" & _
                    CStr(VarType(data(row, column))) & "|" & Text(data(row, column))
            End If
        Next column
    Next row
    ReDim Preserve lines(1 To count)
    Dump = Join(lines, vbCrLf) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: пять форм, три независимых сочетания, все сохранённые НДС.
' formed=Yes вызывает новую первую команду, затем прежнюю команду расчёта.
Public Function RunWorkflowEquivalenceProbe(Optional ByVal formed As String = "No", Optional ByVal selected As String = "") As String
    On Error GoTo Failed
    Dim fixture As Variant, loads As Object, data() As Variant, report As String, message As String
    Dim table As Variant, profiles As Object, values As Variant, row As Long, column As Long, completed As Long
    Setting "General.ExecutionReportEnabled", "No"
    Setting "Plot.Enabled", "No"
    Setting "Load.ReferenceOffsetX", "10": Setting "Load.ReferenceOffsetY", "-7"
    Set profiles = ThisWorkbook.Names("rngCalculationProfiles").RefersToRange
    values = profiles.Value2
    For row = 1 To UBound(values, 1)
        Select Case CStr(values(row, 2))
            Case "Calculation.Strength.DirectState", "Calculation.Strength.Capacity", "Calculation.Crack.Width"
                For column = 3 To UBound(values, 2): values(row, column) = "Yes": Next column
            Case "Calculation.Stability.Enabled"
                For column = 3 To UBound(values, 2): values(row, column) = "No": Next column
        End Select
    Next row
    profiles.Value2 = values
    Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    ReDim data(1 To loads.Rows.Count - 1, 1 To loads.Columns.Count)
    data(1, 1) = "EQ_C": data(1, 2) = 5#: data(1, 3) = 1#: data(1, 4) = -0.5
    data(2, 1) = "EQ_T": data(2, 2) = -2#: data(2, 3) = -0.4: data(2, 4) = 0.7
    data(3, 1) = "EQ_B": data(3, 2) = 0#: data(3, 3) = 1.1: data(3, 4) = 0.6
    For row = 1 To 3: data(row, 5) = "PR2": data(row, 6) = "Auto": Next row
    loads.Offset(1, 0).Resize(UBound(data, 1), UBound(data, 2)).Value2 = data
    For Each fixture In Array("Circle", "Rectangle", "LSection", "RoundedMixed", "HollowOffset")
        If Len(selected) > 0 And selected <> CStr(fixture) Then GoTo NextFixture
        Progress "START|" & CStr(fixture)
        ConfigureShapeFixture CStr(fixture)
        If formed = "Yes" Then message = Application.Run("'" & ThisWorkbook.Name & "'!modWorkbookCalculation.FormSectionForWorkbook", ThisWorkbook)
        message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
        Progress "CALCULATED|" & CStr(fixture)
        For Each table In Array("rngNDMSectionGeometry", "rngNDMSectionContours", "rngNDMSectionAnnotations", _
                "rngNDMSectionProperties", "rngNDMElementResults", "rngNDMMaterialDiagrams")
            report = report & Dump(CStr(fixture), CStr(table), ReadTable(CStr(table)))
        Next table
        report = report & Dump(CStr(fixture), "Summary", ThisWorkbook.Names("rngBatchSummary").RefersToRange.Resize(15, 26).Value2, True)
        completed = completed + 1
        Progress "DUMPED|" & CStr(fixture)
NextFixture:
    Next fixture
    RunWorkflowEquivalenceProbe = report & "TOTAL_WORKFLOW_EQUIVALENCE: passed=" & completed & "; failed=0" & vbCrLf
    Exit Function
Failed:
    RunWorkflowEquivalenceProbe = report & "RUNTIME ERROR: " & Err.Number & "; " & Err.Description & vbCrLf
End Function
