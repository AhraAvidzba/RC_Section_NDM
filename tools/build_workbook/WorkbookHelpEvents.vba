Option Explicit

' ==========================================================================
' События книги: Справка и сохранение связанного файла подробных НДС
' ==========================================================================
' Обработчики управляют справкой и копированием связанного снимка при Save As;
' пользовательский ввод и содержимое расчётного снимка не меняются.
Private mLastWorkbookFolder As String ' Последняя папка книги для последовательных Save As после переноса исходного файла.

' При открытии раскрывает раздел настроек Config и листа «Сечения», остальные сворачивает.
Private Sub Workbook_Open()
    mLastWorkbookFolder = ThisWorkbook.Path
    ApplyHelpDefaults
End Sub

' После Save As переносит файл того же снимка рядом с новой книгой.
' Расчётные значения и другие файлы не меняет; существующий файл проверяет по SHA256.
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    If Not Success Then Exit Sub
    On Error GoTo Failed
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    store.EnsureBesideWorkbook ThisWorkbook, mLastWorkbookFolder
    mLastWorkbookFolder = ThisWorkbook.Path
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Снимок расчёта"
End Sub

' Единое начальное состояние outline-групп для открытия книги.
Private Sub ApplyHelpDefaults()
    On Error GoTo SafeExit
    Dim helpSheet As Worksheet, found As Range
    Set helpSheet = ThisWorkbook.Worksheets("Справка")
    Set found = helpSheet.Cells.Find("Справка по настройкам Config и листа «Сечения»", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext) ' Ищем на всём листе: заголовок объединён A:F.
    If found Is Nothing Then Exit Sub
    helpSheet.Outline.ShowLevels 1
    helpSheet.Rows(found.Row).ShowDetail = True
SafeExit:
End Sub

' Раскрывает только раздел, содержащий адрес выбранной ссылки Подробнее.
Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
    On Error GoTo SafeExit
    Dim addressText As String, helpSheet As Worksheet, targetCell As Range
    If Not TryGetHelpAddress(Target.SubAddress, addressText) Then Exit Sub
    Set helpSheet = ThisWorkbook.Worksheets("Справка")
    Set targetCell = helpSheet.Range(addressText)
    RevealHelpOutlineGroup helpSheet, targetCell.Row
    Application.Goto targetCell, True
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.ScrollRow = targetCell.Row
        ActiveWindow.ScrollColumn = 1
    End If
SafeExit:
End Sub

' Разбирает внутреннюю ссылку с кавычками или без них вокруг имени Справка.
Private Function TryGetHelpAddress(ByVal subAddress As String, ByRef addressText As String) As Boolean
    Dim bangPos As Long, sheetName As String
    bangPos = InStrRev(subAddress, "!")
    If bangPos <= 1 Or bangPos >= Len(subAddress) Then Exit Function
    sheetName = Replace(Left$(subAddress, bangPos - 1), "'", vbNullString)
    If StrComp(sheetName, "Справка", vbTextCompare) <> 0 Then Exit Function
    addressText = Mid$(subAddress, bangPos + 1)
    TryGetHelpAddress = Len(addressText) > 0
End Function

' Находит верхний заголовок единой группы, не раскрывая соседние разделы.
Private Sub RevealHelpOutlineGroup(ByVal helpSheet As Worksheet, ByVal targetRow As Long)
    On Error GoTo SafeExit
    Dim firstRow As Long, lastRow As Long, level As Long
    level = helpSheet.Rows(targetRow).OutlineLevel
    If level <= 1 Then
        If helpSheet.Rows(targetRow + 1).OutlineLevel > 1 Then helpSheet.Rows(targetRow).ShowDetail = True
        Exit Sub
    End If
    firstRow = targetRow
    Do While firstRow > 1
        If helpSheet.Rows(firstRow - 1).OutlineLevel < level Then Exit Do
        firstRow = firstRow - 1
    Loop
    lastRow = helpSheet.Cells(helpSheet.Rows.Count, 1).End(xlUp).Row
    If firstRow > 1 Then helpSheet.Rows(firstRow - 1).ShowDetail = True
SafeExit:
End Sub
