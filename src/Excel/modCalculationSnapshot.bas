Attribute VB_Name = "modCalculationSnapshot"
Option Explicit

' ==========================================================================
' Общая загрузка и восстановление результатов без повторного расчёта
' ==========================================================================
' После расчёта вывод использует готовые данные в памяти. Кнопка в File
' читает текущий выбранный файл расчёта либо геометрии; Excel читает только Results.
' Кратковременный контекст чтения принадлежит одной операции и одной книге.
Private mReadSnapshot As CCalculationSnapshot ' Неизменяемый набор только на время одной операции чтения.
Private mReadWorkbook As Object ' Книга, которой принадлежит текущий контекст.
Private mReadDepth As Long ' Вложенные чтения схемы и экспорта не уничтожают внешний контекст.
Private mReadSourceResolved As Boolean ' Выбор Excel/File проверяется один раз на операцию, в том числе при отсутствии файла в Excel.

' Открывает короткую область чтения; между пользовательскими действиями кэша нет.
Public Sub BeginCalculationSnapshotRead(ByVal workbook As Object, Optional ByVal snapshot As CCalculationSnapshot = Nothing)
    If mReadDepth = 0 Then
        If Not snapshot Is Nothing Then snapshot.Validate
        Set mReadWorkbook = workbook: Set mReadSnapshot = snapshot
        mReadSourceResolved = Not snapshot Is Nothing ' Готовый набор принадлежит текущему выводу в обоих режимах; ручное чтение без него по-прежнему выбирает Config.
    ElseIf Not workbook Is mReadWorkbook Then
        Err.Raise 5, "modCalculationSnapshot", "Вложенное чтение относится к другой книге."
    ElseIf Not snapshot Is Nothing Then
        snapshot.Validate
        Set mReadSnapshot = snapshot
        mReadSourceResolved = True
    End If
    mReadDepth = mReadDepth + 1
End Sub

' Освобождает большие массивы и ссылку на книгу после успеха или ошибки.
Public Sub EndCalculationSnapshotRead()
    If mReadDepth > 0 Then mReadDepth = mReadDepth - 1
    If mReadDepth = 0 Then Set mReadSnapshot = Nothing: Set mReadWorkbook = Nothing: mReadSourceResolved = False
End Sub

' Выбирает источник только по текущему Excel/File. File сразу читает указанное имя,
' Excel не обращается к файлу независимо от прежней привязки или смены имени.
Public Function CurrentCalculationSnapshot(ByVal workbook As Object) As CCalculationSnapshot
    If workbook Is mReadWorkbook Then
        If mReadSourceResolved Then Set CurrentCalculationSnapshot = mReadSnapshot: Exit Function
    End If
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim snapshot As CCalculationSnapshot
    If store.StorageMode(workbook) = "File" Then
        If workbook Is mReadWorkbook Then Set snapshot = mReadSnapshot ' Уже проверенные данные текущего сохранения/загрузки не читаются повторно в той же операции.
        If snapshot Is Nothing Then Set snapshot = store.LoadFile(workbook)
    End If
    If mReadDepth > 0 And workbook Is mReadWorkbook Then Set mReadSnapshot = snapshot: mReadSourceResolved = True
    Set CurrentCalculationSnapshot = snapshot
End Function

' Выводит готовый расчёт из памяти либо восстанавливает полный снимок File.
' Ручное обновление Excel читает схему с листа и не восстанавливает верхние сводки.
Public Sub UpdateSectionResultsForWorkbook(ByVal workbook As Object, Optional ByVal prepared As CCalculationSnapshot = Nothing, _
        Optional ByRef plotWarning As String = "", Optional ByRef plotHasWarning As Boolean = False)
    plotHasWarning = False
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
        Else
            store.ClearReference workbook ' Источник Excel — только лист; прежняя служебная привязка ничего не подгружает.
            UpdateSectionPlotForWorkbook workbook
            guard.Restore True
    CompleteWorkbookDrawing workbook
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
    writer.RestoreSnapshotBase workbook, snapshot, fullFormatting, mode = "Excel"
    If snapshot.IsGeometryPreview Then
        summary.ClearSummary workbook, fullFormatting
    Else
        summary.RestoreSummary workbook, snapshot, fullFormatting
    End If
    If mode = "Excel" Then
        store.ClearReference workbook
    Else
        store.LinkFile workbook
    End If
    If snapshot.IsGeometryPreview Then
        UpdateSectionGeometryPreviewForWorkbook workbook
    ElseIf prepared Is Nothing Then
        UpdateSectionPlotForWorkbook workbook, False
    Else
        plotWarning = TryUpdateCalculatedSnapshotPlot(workbook, snapshot, plotHasWarning)
    End If
    EndCalculationSnapshotRead
    entered = False
    guard.Restore True
    CompleteWorkbookDrawing workbook
    Exit Sub
Failed:
    Dim errorNumber As Long, reason As String, source As String
    errorNumber = Err.Number: reason = ResultOutputErrorDescription(Err.Description): source = Err.source
    If entered Then EndCalculationSnapshotRead
    guard.Restore
    Err.Raise errorNumber, source, reason
End Sub

' Проверяет вместимость готовых таблиц до очистки; число LC не берётся из текущего Config.
Private Sub ValidateRestoredSnapshotLayout(ByVal workbook As Object, ByVal snapshot As CCalculationSnapshot, ByVal includeElements As Boolean)
    Dim key As Variant, values As Variant, anchor As Object, rowCount As Long
    Set anchor = workbook.Names.Item("rngNDMElementGeometry").RefersToRange ' В старом payload её может не быть; текущая книга всё равно выводит шапку до любых очисток.
    If Not ResultRowsFitSheet(anchor.Row, 1, anchor.Worksheet.Rows.Count) Then Err.Raise 5, "ValidateRestoredSnapshotLayout", "Не помещается шапка rngNDMElementGeometry."
    For Each key In snapshot.BaseKeys()
        Set anchor = workbook.Names.Item(CStr(key)).RefersToRange
        rowCount = snapshot.TableRowCount(CStr(key))
        If Not includeElements Then
            If CStr(key) = "rngNDMElementResults" Or CStr(key) = "rngNDMElementGeometry" Or CStr(key) = "rngNDMSectionAnnotations" Then rowCount = 1
        End If
        If includeElements And Not snapshot.IsGeometryPreview Then
            If CStr(key) = "rngNDMSectionProperties" Then rowCount = rowCount + snapshot.TableRowCount("LoadCase.Properties") - 1
            If CStr(key) = "rngNDMSectionAnnotations" Then rowCount = rowCount + snapshot.TableRowCount("LoadCase.Annotations") - 1
        End If
        If Not ResultRowsFitSheet(anchor.Row, rowCount, anchor.Worksheet.Rows.Count) Then
            Dim reason As String: reason = ResultRowLimitDescription(CStr(key), anchor.Row, rowCount, anchor.Worksheet.Rows.Count)
            If CStr(key) = "rngNDMElementResults" Or CStr(key) = "rngNDMElementGeometry" Then
                If RestoredSnapshotMandatoryRowsFit(workbook, snapshot) Then reason = reason & vbCrLf & "Что сделать:" & vbCrLf & "1) " & ExternalSnapshotStorageAdvice(workbook) & " Повторите «Обновить схему и результаты»."
            End If
            Err.Raise 5, "ValidateRestoredSnapshotLayout", reason
        End If
    Next key
    If snapshot.IsGeometryPreview Then Exit Sub
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

' Проверяет таблицы, остающиеся на листе при File, без копирования матриц.
' Используется только при отказе вместимости таблицы состояний: File нельзя
' предлагать, если общая геометрия или готовые сводки тоже не помещаются.
Private Function RestoredSnapshotMandatoryRowsFit(ByVal workbook As Object, ByVal snapshot As CCalculationSnapshot) As Boolean
    Dim key As Variant, anchor As Object
    For Each key In snapshot.BaseKeys()
        If CStr(key) <> "rngNDMElementResults" And CStr(key) <> "rngNDMElementGeometry" And CStr(key) <> "rngNDMSectionAnnotations" Then
            Set anchor = workbook.Names.Item(CStr(key)).RefersToRange
            If Not ResultRowsFitSheet(anchor.Row, snapshot.TableRowCount(CStr(key)), anchor.Worksheet.Rows.Count) Then Exit Function
        End If
    Next key
    If snapshot.IsGeometryPreview Then RestoredSnapshotMandatoryRowsFit = True: Exit Function
    Dim names As Variant, keys As Variant, headerRows As Variant, i As Long
    names = Array("rngBatchSummary", "rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor")
    keys = Array("Summary", "Strength.Data", "Cracks.Data", "Stability.Data")
    headerRows = Array(0, 4, 4, 5)
    For i = 0 To 3
        Set anchor = workbook.Names.Item(CStr(names(i))).RefersToRange
        If Not ResultRowsFitSheet(anchor.Row - CLng(headerRows(i)), CDbl(snapshot.TableRowCount(CStr(keys(i)))) + CLng(headerRows(i)), anchor.Worksheet.Rows.Count) Then Exit Function
    Next i
    RestoredSnapshotMandatoryRowsFit = True
End Function

' Проверяет фактическую вместимость листа с учётом первой строки, без переполнения Long.
' Предел передаётся из Worksheet.Rows.Count: .xls и .xlsm имеют разную высоту листа.
Public Function ResultRowsFitSheet(ByVal firstRow As Long, ByVal rowCount As Double, ByVal rowLimit As Long) As Boolean
    If firstRow < 1 Or rowCount < 1# Or rowLimit < firstRow Then Exit Function
    ResultRowsFitSheet = rowCount <= CDbl(rowLimit) - CDbl(firstRow) + 1#
End Function

' Называет потребное число строк и фактический предел текущего листа, включая смещение якоря.
Public Function ResultRowLimitDescription(ByVal tableName As String, ByVal firstRow As Long, ByVal rowCount As Double, ByVal rowLimit As Long) As String
    ResultRowLimitDescription = tableName & " требует " & Format$(rowCount, "0") & " строк от строки " & CStr(firstRow) & _
        " (последняя строка " & Format$(CDbl(firstRow) + rowCount - 1#, "0") & "). " & _
        "Этот лист Excel поддерживает " & CStr(rowLimit) & " строк; ниже начала диапазона доступно " & CStr(rowLimit - firstRow + 1) & "."
End Function

' Предлагает File с реальным адресом настройки; выбор пользователя автоматически не изменяет.
Public Function ExternalSnapshotStorageAdvice(ByVal workbook As Object) As String
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
    settings.LoadFromRange workbook.Names.Item("rngSystemSettings").RefersToRange
    ExternalSnapshotStorageAdvice = "На Config установите Results.SnapshotStorage = File (" & settings.InputLocation("Results.SnapshotStorage") & _
        "), чтобы сохранять полный снимок во внешнем файле .ndm: для данных снимка на листе Excel недостаточно строк."
End Function

' Объясняет подтверждённый отказ Excel из-за объединённых ячеек, сохраняя прочие причины без подмены.
' Одного кода 1004 недостаточно: он используется Excel для множества разных ошибок.
Public Function ResultOutputErrorDescription(ByVal original As String) As String
    ResultOutputErrorDescription = original
    Dim normalized As String: normalized = LCase$(Replace$(original, "ё", "е"))
    If (InStr(normalized, "объединен") > 0 And InStr(normalized, "ячей") > 0) Or InStr(normalized, "merged cell") > 0 Then
        ResultOutputErrorDescription = "Не удалось вывести результаты на лист Results: область вывода попадает в объединённые ячейки. " & _
            "Снимите объединение в месте вывода или перенесите именованный диапазон в свободную необъединённую область, затем повторите действие."
    End If
End Function

' Сохраняет полный снимок только в File; оба режима выводят результаты из памяти.
' Excel оставляет обычные читаемые таблицы без второго архива подробных сводок.
Public Function SaveAndDisplayCalculationSnapshot(ByVal workbook As Object, ByVal snapshot As CCalculationSnapshot, _
        Optional ByRef plotHasWarning As Boolean = False) As String
    plotHasWarning = False
    snapshot.Validate
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    ValidateRestoredSnapshotLayout workbook, snapshot, store.StorageMode(workbook) = "Excel"
    If store.StorageMode(workbook) = "File" Then
        store.SaveFile workbook, snapshot
    End If
    Dim warning As String
    UpdateSectionResultsForWorkbook workbook, snapshot, warning, plotHasWarning
    SaveAndDisplayCalculationSnapshot = warning
End Function

' Ошибка оформления не отменяет сохранённый расчёт и заполненные подробные таблицы.
' Ручное обновление возвращает исходную ошибку; автоматическое завершение — предупреждение.
Private Function TryUpdateCalculatedSnapshotPlot(ByVal workbook As Object, ByVal snapshot As CCalculationSnapshot, _
        ByRef plotHasWarning As Boolean) As String
    plotHasWarning = False
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
    plotHasWarning = True
    TryUpdateCalculatedSnapshotPlot = "Предупреждение:" & vbCrLf & "Схема не обновлена. Расчётные результаты сохранены." & vbCrLf & _
        Err.Description & vbCrLf & "Исправьте настройки оформления и нажмите «Обновить схему и результаты». Повторный расчёт не требуется."
End Function

' Выбирает предупреждение по числу строк и фактической вместимости листа.
' mandatoryRows содержит записи (имя диапазона, число строк, первая строка)
' всех таблиц, остающихся на Results при File. RowLimit запрещает предложение
' File: переполнение обязательной геометрии режим хранения не устраняет.
Public Function CalculationSnapshotOutputPolicy(ByVal storageMode As String, ByVal estimatedRows As Double, _
        ByRef mandatoryRows As Variant, ByVal rowLimit As Long, Optional ByVal estimatedBoundaryRows As Double = 0#) As String
    Dim i As Long, entry As Variant, largeMandatory As Boolean
    If IsArray(mandatoryRows) Then
        For i = LBound(mandatoryRows) To UBound(mandatoryRows)
            entry = mandatoryRows(i)
            If Not ResultRowsFitSheet(CLng(entry(2)), CDbl(entry(1)), rowLimit) Then
                CalculationSnapshotOutputPolicy = "RowLimit"
                Exit Function
            End If
            If CDbl(entry(1)) > 60000# Then largeMandatory = True
        Next i
    End If
    If storageMode = "Excel" And (estimatedRows > 60000# Or estimatedBoundaryRows > 60000#) Then CalculationSnapshotOutputPolicy = "File"
    If largeMandatory Then
        If Len(CalculationSnapshotOutputPolicy) > 0 Then
            CalculationSnapshotOutputPolicy = "FileAndMandatory"
        Else
            CalculationSnapshotOutputPolicy = "Mandatory"
        End If
    End If
End Function

' Предупреждает о крупных обязательных таблицах в Excel и File, независимо
' от способа формирования сетки. Переключение в File предлагается только
' для больших состояний или собственных границ и при вместимости остальных таблиц Results.
Public Sub OfferExternalCalculationSnapshot(ByVal workbook As Object, ByVal estimatedRows As Double, ByVal interactive As Boolean, _
        Optional ByVal mandatoryRows As Variant, Optional ByVal estimatedBoundaryRows As Double = 0#)
    If Not interactive Then Exit Sub
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim policy As String: policy = CalculationSnapshotOutputPolicy(store.StorageMode(workbook), estimatedRows, _
        mandatoryRows, workbook.Worksheets.Item("Results").Rows.Count, estimatedBoundaryRows)
    If Len(policy) = 0 Or policy = "RowLimit" Then Exit Sub ' Отказ по пределу листа сообщает общая предварительная проверка до расчёта.
    Dim message As String, i As Long, entry As Variant, table As Object
    If policy = "Mandatory" Or policy = "FileAndMandatory" Then
        message = "На Results предстоит вывести большие обязательные таблицы."
        For i = LBound(mandatoryRows) To UBound(mandatoryRows)
            entry = mandatoryRows(i)
            If CDbl(entry(1)) > 60000# Then
                Set table = workbook.Names.Item(CStr(entry(0))).RefersToRange
                message = message & vbCrLf & CStr(entry(0)) & " (" & table.Parent.Name & "!" & _
                    table.Address(False, False) & "): около " & Format$(CDbl(entry(1)), "0") & " строк."
            End If
        Next i
        message = message & vbCrLf & vbCrLf & "Эти таблицы сохраняются на листе и при Results.SnapshotStorage = File. " & _
            "Переключение снимка в файл не уменьшает число строк общей геометрии, общих контуров, характеристик и сводок."
    End If
    If policy = "Mandatory" Then
        ShowWorkbookMessage message, vbExclamation, "Большой объём Results"
        Exit Sub
    End If
    If Len(message) > 0 Then message = message & vbCrLf & vbCrLf
    message = message & "Предварительная оценка подробных таблиц:" & vbCrLf & _
        "Состояния элементов: около " & Format$(estimatedRows, "0") & " строк." & vbCrLf & _
        "Фактические границы элементов: около " & Format$(estimatedBoundaryRows, "0") & " строк." & vbCrLf & vbCrLf & _
        "Внешний файл .ndm сохранит полные состояния и фактические границы элементов без их вывода на лист. " & _
        "Общая геометрия, общие контуры, использованные диаграммы и подробные сводки останутся на Results." & vbCrLf & vbCrLf & _
        "Переключить Results.SnapshotStorage на File?" & vbCrLf & _
        "Да — переключить и продолжить. Нет — продолжить с хранением в Excel."
    Dim answer As VbMsgBoxResult: answer = MsgBox(message, vbQuestion Or vbYesNo, "Хранение большого снимка")
    If answer <> vbYes Then Exit Sub
    Dim values As Variant
    Set table = workbook.Names.Item("rngSystemSettings").RefersToRange: values = table.Value2
    For i = 1 To UBound(values, 1)
        If CStr(values(i, 1)) = "Results.SnapshotStorage" Then table.Cells(i, 2).Value2 = "File": Exit Sub
    Next i
    Err.Raise 5, , "В rngSystemSettings отсутствует Results.SnapshotStorage."
End Sub

' Удаляет привязку готового расчёта перед обычной записью preview или тестового writer.
Public Sub ClearPublishedCalculationSnapshot(ByVal workbook As Object)
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    store.ClearReference workbook
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Доказывает освобождение большого снимка и книги после законченной операции.
Public Function TestCalculationSnapshotReadReleased() As Boolean
    TestCalculationSnapshotReadReleased = (mReadDepth = 0 And mReadSnapshot Is Nothing And mReadWorkbook Is Nothing)
End Function
