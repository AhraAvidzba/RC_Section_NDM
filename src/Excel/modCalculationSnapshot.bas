Attribute VB_Name = "modCalculationSnapshot"
Option Explicit

' ==========================================================================
' Общая загрузка и восстановление результатов без повторного расчёта
' ==========================================================================
' Расчёт сохраняет полный снимок, затем вызывает тот же путь, что кнопка.
' Кратковременный контекст чтения принадлежит одной операции и одной книге.
Private mReadSnapshot As CCalculationSnapshot ' Неизменяемый набор только на время одной операции чтения.
Private mReadWorkbook As Object ' Книга, которой принадлежит текущий контекст.
Private mReadDepth As Long ' Вложенные чтения схемы и экспорта не уничтожают внешний контекст.

' Открывает короткую область чтения; между пользовательскими действиями кэша нет.
Public Sub BeginCalculationSnapshotRead(ByVal workbook As Object, Optional ByVal snapshot As CCalculationSnapshot = Nothing)
    If mReadDepth = 0 Then
        Set mReadWorkbook = workbook: Set mReadSnapshot = snapshot
    ElseIf Not workbook Is mReadWorkbook Then
        Err.Raise 5, "modCalculationSnapshot", "Вложенное чтение относится к другой книге."
    ElseIf Not snapshot Is Nothing Then
        Set mReadSnapshot = snapshot
    End If
    mReadDepth = mReadDepth + 1
End Sub

' Освобождает большие массивы и ссылку на книгу после успеха или ошибки.
Public Sub EndCalculationSnapshotRead()
    If mReadDepth > 0 Then mReadDepth = mReadDepth - 1
    If mReadDepth = 0 Then Set mReadSnapshot = Nothing: Set mReadWorkbook = Nothing
End Sub

' Читает именно показанный расчёт; новый выбор имени применяется кнопкой загрузки.
' Preview и прежний обычный writer не обязаны иметь полный расчётный снимок.
Public Function CurrentCalculationSnapshot(ByVal workbook As Object) As CCalculationSnapshot
    If workbook Is mReadWorkbook Then
        If Not mReadSnapshot Is Nothing Then Set CurrentCalculationSnapshot = mReadSnapshot: Exit Function
    End If
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim snapshot As CCalculationSnapshot
    If store.HasFileReference(workbook) Then
        Set snapshot = store.LoadLinkedFile(workbook)
    ElseIf store.HasSheetSnapshot(workbook) Then
        Set snapshot = store.LoadSheet(workbook)
    Else
        Exit Function
    End If
    If mReadDepth > 0 And workbook Is mReadWorkbook Then Set mReadSnapshot = snapshot
    Set CurrentCalculationSnapshot = snapshot
End Function

' Восстанавливает общие данные сечения и все подробные сводки выбранного снимка.
' Файл читается по имени Config, а не по старой ссылке. Solver и импорт не вызываются.
Public Sub UpdateSectionResultsForWorkbook(ByVal workbook As Object, Optional ByVal prepared As CCalculationSnapshot = Nothing, _
        Optional ByVal sheetAlreadySaved As Boolean = False, Optional ByRef plotWarning As String = "")
    Dim guard As CExcelAppStateGuard: Set guard = New CExcelAppStateGuard
    Dim entered As Boolean
    On Error GoTo Failed
    guard.Enter workbook.Application
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim mode As String: mode = store.StorageMode(workbook)
    Dim snapshot As CCalculationSnapshot: Set snapshot = prepared
    If snapshot Is Nothing Then
        If mode = "File" Then
            Set snapshot = store.LoadFile(workbook)
        ElseIf store.HasSheetSnapshot(workbook) Then
            Set snapshot = store.LoadSheet(workbook)
        ElseIf store.HasFileReference(workbook) Then
            Set snapshot = store.LoadLinkedFile(workbook)
        Else
            UpdateSectionPlotForWorkbook workbook
            guard.Restore True
            Exit Sub
        End If
    End If
    snapshot.Validate
    ValidateRestoredSnapshotLayout workbook, snapshot, mode = "Excel"
    BeginCalculationSnapshotRead workbook, snapshot: entered = True
    Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    Dim summary As CBatchResultWriter: Set summary = New CBatchResultWriter
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook workbook
    Dim fullFormatting As Boolean: fullFormatting = settings.GetRequiredBoolean("General.ReformatResultsEnabled")
    If Not sheetAlreadySaved Then writer.RestoreSnapshotBase workbook, snapshot, fullFormatting, mode = "Excel"
    summary.RestoreSummary workbook, snapshot, fullFormatting
    If mode = "Excel" Then
        store.ClearReference workbook
        If Not sheetAlreadySaved Then store.SaveSheetExtra workbook, snapshot
    Else
        store.ClearSheetExtra workbook
        store.LinkFile workbook
    End If
    If prepared Is Nothing Then
        UpdateSectionPlotForWorkbook workbook, False
    Else
        plotWarning = TryUpdateCalculatedSnapshotPlot(workbook, snapshot)
    End If
    EndCalculationSnapshotRead
    entered = False
    guard.Restore True
    Exit Sub
Failed:
    Dim errorNumber As Long, reason As String, source As String
    errorNumber = Err.Number: reason = Err.Description: source = Err.source
    If entered Then EndCalculationSnapshotRead
    guard.Restore
    Err.Raise errorNumber, source, reason
End Sub

' Проверяет вместимость готовых таблиц до очистки; число LC не берётся из текущего Config.
Private Sub ValidateRestoredSnapshotLayout(ByVal workbook As Object, ByVal snapshot As CCalculationSnapshot, ByVal includeElements As Boolean)
    Dim key As Variant, values As Variant, anchor As Object
    For Each key In snapshot.BaseKeys()
        If CStr(key) <> "rngNDMElementResults" Or includeElements Then
            Set anchor = workbook.Names.Item(CStr(key)).RefersToRange
            If snapshot.TableRowCount(CStr(key)) > anchor.Worksheet.Rows.Count - anchor.Row + 1 Then Err.Raise 5, , _
                "Таблица " & CStr(key) & " не помещается на Results. Для большого снимка выберите Results.SnapshotStorage = File."
        End If
    Next key
    Dim names As Variant, keys As Variant, nextNames As Variant, headers As Variant, i As Long
    names = Array("rngBatchSummary", "rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor")
    keys = Array("Summary", "Strength.Data", "Cracks.Data", "Stability.Data")
    nextNames = Array("rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor", "rngNDMSectionGeometry")
    headers = Array(4, 4, 5, 2)
    For i = 0 To 3
        Set anchor = workbook.Names.Item(CStr(names(i))).RefersToRange: values = snapshot.Table(CStr(keys(i)))
        Dim nextAnchor As Object: Set nextAnchor = workbook.Names.Item(CStr(nextNames(i))).RefersToRange
        If anchor.Row + UBound(values, 1) + CLng(headers(i)) + 1 >= nextAnchor.Row Then Err.Raise 5, , _
            "В снимке больше строк, чем в текущей раскладке Results: " & CStr(names(i)) & " пересекает " & CStr(nextNames(i)) & ". Увеличьте расстояние между именованными диапазонами."
    Next i
End Sub

' Сохраняет снимок перед общим обновлением. В режиме Excel прежние общие
' таблицы и один LC-блок вместе составляют тот же набор, что внешний файл.
Public Function SaveAndDisplayCalculationSnapshot(ByVal workbook As Object, ByVal snapshot As CCalculationSnapshot) As String
    snapshot.Validate
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim onSheet As Boolean: onSheet = (store.StorageMode(workbook) = "Excel")
    If Not onSheet Then
        store.SaveFile workbook, snapshot
    Else
        Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
        Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook workbook
        writer.RestoreSnapshotBase workbook, snapshot, settings.GetRequiredBoolean("General.ReformatResultsEnabled")
        store.ClearReference workbook
        store.SaveSheetExtra workbook, snapshot
    End If
    Dim warning As String
    UpdateSectionResultsForWorkbook workbook, snapshot, onSheet, warning
    SaveAndDisplayCalculationSnapshot = warning
End Function

' Ошибка оформления не отменяет сохранённый расчёт и заполненные подробные таблицы.
' Ручное обновление возвращает исходную ошибку; автоматическое завершение — предупреждение.
Private Function TryUpdateCalculatedSnapshotPlot(ByVal workbook As Object, ByVal snapshot As CCalculationSnapshot) As String
    On Error GoTo Failed
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook workbook
    If Not settings.GetRequiredBoolean("Plot.Enabled") Then Exit Function
    If Not settings.GetRequiredBoolean("Plot.AutoUpdateAfterCalculation") Then Exit Function
    If snapshot.TableRowCount("rngNDMElementResults") <= 1 Then
        Dim plotter As CSectionPlotter: Set plotter = New CSectionPlotter
        plotter.ClearExisting workbook, "Схема не обновлена: нет доступных расчётных состояний LC."
        Exit Function
    End If
    UpdateSectionPlotForWorkbook workbook, False
    Exit Function
Failed:
    TryUpdateCalculatedSnapshotPlot = "Предупреждение:" & vbCrLf & "Схема не обновлена. Снимок и результаты сохранены." & vbCrLf & _
        Err.Description & vbCrLf & "Исправьте настройки оформления и нажмите «Обновить схему и результаты». Повторный расчёт не требуется."
End Function

' Предлагает более подходящее хранение большого снимка, сохраняя выбор Excel при отказе.
' Оценка не является ограничением расчёта и не выдаётся за точное число конечных строк.
Public Sub OfferExternalCalculationSnapshot(ByVal workbook As Object, ByVal estimatedRows As Double, ByVal interactive As Boolean)
    If Not interactive Or estimatedRows <= 60000# Then Exit Sub
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    If store.StorageMode(workbook) <> "Excel" Then Exit Sub
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Предварительная оценка объёма снимка — около " & Format$(estimatedRows, "#,##0") & _
        " строк состояний элементов." & vbCrLf & vbCrLf & _
        "Для такого объёма рекомендуется внешний файл: книга будет меньше, сохранение и загрузка большого снимка потребуют меньше служебных данных на листе. " & _
        "Сетка, геометрические характеристики, использованные диаграммы и подробные сводки останутся доступны на Results." & vbCrLf & vbCrLf & _
        "Переключить Results.SnapshotStorage на File?" & vbCrLf & _
        "Да — переключить и продолжить. Нет — продолжить расчёт с хранением в Excel.", vbQuestion Or vbYesNo, "Хранение большого снимка")
    If answer <> vbYes Then Exit Sub
    Dim table As Object, values As Variant, i As Long
    Set table = workbook.Names.Item("rngSystemSettings").RefersToRange: values = table.Value2
    For i = 1 To UBound(values, 1)
        If CStr(values(i, 1)) = "Results.SnapshotStorage" Then table.Cells(i, 2).Value2 = "File": Exit Sub
    Next i
    Err.Raise 5, , "В rngSystemSettings отсутствует Results.SnapshotStorage."
End Sub

' Удаляет привязку готового расчёта перед обычной записью preview или тестового writer.
' Отсутствующий новый якорь в изолированной тестовой книге не создаётся автоматически.
Public Sub ClearPublishedCalculationSnapshot(ByVal workbook As Object)
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    store.ClearReference workbook
    Dim target As Object
    On Error Resume Next
    Set target = workbook.Names.Item("rngNDMLoadCaseResults").RefersToRange
    On Error GoTo 0
    If Not target Is Nothing Then store.ClearSheetExtra workbook
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Доказывает освобождение большого снимка и книги после законченной операции.
Public Function TestCalculationSnapshotReadReleased() As Boolean
    TestCalculationSnapshotReadReleased = (mReadDepth = 0 And mReadSnapshot Is Nothing And mReadWorkbook Is Nothing)
End Function
