VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBatchResultWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Сводка пакетного расчета на листе Results
' ==========================================================================
' Writer формирует верхний диапазон rngBatchSummary как якорную компактную
' сводку по всем сочетаниям. В ней остаются только выбранное Worst-сочетание,
' статусы проверок и минимальные коэффициенты запаса; подробные данные
' прочности, трещин и устойчивости пишутся отдельными writer-ами ниже.
' Расчетные алгоритмы здесь не запускаются: класс только берет готовые данные
' CBatchSectionCalculator и записывает их на лист одним массивом.

Private Const TARGET_COLUMNS As Long = 15 ' Ширина верхней сводки без детальных блоков трещин и устойчивости.
Private Const LEGACY_SUMMARY_CLEAR_COLUMNS As Long = 34 ' Чистит старую широкую сводку и правый словарь статусов при первом запуске новой версии.
Private Const META_ROWS As Long = 4 ' Верхние строки с общими сведениями о запуске.
Private Const GROUP_HEADER_ROW As Long = 6 ' Первая строка шестистрочной шапки основной таблицы.
Private Const SUBGROUP_HEADER_ROW As Long = 7 ' Строка внутренних подзаголовков расчетных блоков.
Private Const COLUMN_HEADER_ROW As Long = 8 ' Верх строки с обозначениями конкретных параметров.
Private Const COLUMN_HEADER_LAST_ROW As Long = 10 ' Низ объединенных ячеек с обозначениями параметров.
Private Const COMMENT_HEADER_ROW As Long = 11 ' Последняя строка шапки: короткие комментарии к столбцам.
Private Const FIRST_DATA_ROW As Long = 12 ' Первая строка сочетаний.

Private Const SUM_COL_ID As Long = 1 ' CombinationID в новой компактной сводке.
Private Const SUM_COL_COMMENT As Long = 2 ' Пользовательский комментарий LC.
Private Const SUM_COL_OVERALL_STATUS As Long = 3 ' Итоговый статус LC по всем включенным проверкам.
Private Const SUM_COL_STATUS_STRAIN As Long = 5 ' Статус прямого НДС по модели прочности.
Private Const SUM_COL_STATUS_CAPACITY As Long = 6 ' Статус проверки несущей способности.
Private Const SUM_COL_STATUS_CRACK As Long = 7 ' Статус нормальных трещин.
Private Const SUM_COL_STATUS_LONG_CRACK As Long = 8 ' Статус продольных трещин.
Private Const SUM_COL_STATUS_STABILITY As Long = 9 ' Статус устойчивости.
Private Const SUM_COL_RESERVE_STRAIN As Long = 11 ' Запас по деформациям strength-модели.
Private Const SUM_COL_RESERVE_CAPACITY As Long = 12 ' Запас по предельным усилиям.
Private Const SUM_COL_RESERVE_CRACK As Long = 13 ' Запас по нормальным трещинам.
Private Const SUM_COL_RESERVE_LONG_CRACK As Long = 14 ' Запас по продольным трещинам.
Private Const SUM_COL_RESERVE_STABILITY As Long = 15 ' Запас по устойчивости.

Private Const DEFAULT_SUMMARY_ROWS As Long = 31 ' Штатная высота шаблона: 11 строк до данных и 20 строк LC.
Private Const RESULTS_COLUMN_WIDTH As Double = 10# ' Единая ширина всех столбцов Results.

Private Const HEADER_COLOR As Long = &HD9D9D9 ' Светло-серый фон заголовков.
Private Const SAFETY_COLOR As Long = &HE7E6E6 ' Серый фон столбцов коэффициентов запаса.
Private Const FAILED_ROW_COLOR As Long = &HCEC7FF ' Светло-красный фон LC, где расчет не прошел или не сошелся.
Private Const FAILED_FONT_COLOR As Long = &H9C0006 ' Темно-красный текст для строк FAIL/NumFail/InputErr.

' Возвращает высоту основной таблицы Summary для фактического количества LC.
Public Function RequiredSummaryRows(ByVal combinationCount As Long) As Long
    If combinationCount < 0 Then combinationCount = 0
    RequiredSummaryRows = FIRST_DATA_ROW - 1 + combinationCount
    If RequiredSummaryRows < COLUMN_HEADER_ROW Then RequiredSummaryRows = COLUMN_HEADER_ROW
    If RequiredSummaryRows < DEFAULT_SUMMARY_ROWS Then RequiredSummaryRows = DEFAULT_SUMMARY_ROWS
End Function

' Возвращает ширину собственно rngBatchSummary.
Public Function RequiredSummaryColumns() As Long
    RequiredSummaryColumns = TARGET_COLUMNS
End Function

' Возвращает полную ширину, которую writer затрагивает на Results.
Public Function RequiredSummaryOutputColumns() As Long
    RequiredSummaryOutputColumns = TARGET_COLUMNS
End Function

' Возвращает полную высоту, которую занимает компактная сводка.
Public Function RequiredSummaryOutputRows(ByVal combinationCount As Long) As Long
    RequiredSummaryOutputRows = RequiredSummaryRows(combinationCount)
End Function

' Записывает rngBatchSummary одним массивом и оформляет блоки.
' section нужен для геометрических величин h/x: в блоке деформаций они
' относятся к фактическому LC, а в блоке предельного момента - только к
' найденному capacity-состоянию. Если section не передан, эти поля пустые.
Public Sub WriteSummary(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, _
        Optional ByVal units As CUnitSystem = Nothing, Optional ByVal section As CSectionModel = Nothing)
    On Error GoTo Failed

    Dim stage As String
    stage = "ValidateArguments"
    If workbook Is Nothing Then Err.Raise vbObjectError + 3970, "CBatchResultWriter", "Книга Excel не передана."
    If batch Is Nothing Then Err.Raise vbObjectError + 3971, "CBatchResultWriter", "Пакетный расчетчик не передан."

    Dim summaryRows As Long
    summaryRows = RequiredSummaryRows(batch.Count)

    stage = "PrepareSummaryTarget"
    PrepareSummaryTarget SummaryClearRange(workbook)

    stage = "SummaryTargetRange"
    Dim target As Object
    Set target = SummaryTargetRange(workbook, summaryRows)

    stage = "AllocateData"
    Dim data() As Variant
    ReDim data(1 To summaryRows, 1 To TARGET_COLUMNS)

    stage = "PutMetaRows"
    PutMetaRows data, batch, units
    stage = "PutGroupHeaders"
    PutGroupHeaders data
    stage = "PutSubgroupHeaders"
    PutSubgroupHeaders data
    stage = "PutColumnHeaders"
    PutColumnHeaders data, units

    Dim i As Long
    For i = 1 To batch.Count
        stage = "PutCombination " & CStr(i)
        PutCombination data, FIRST_DATA_ROW + i - 1, batch, i, units, section
    Next i

    stage = "WriteSummaryBody"
    WriteSummaryBody target, data
    stage = "FormatSummary"
    FormatSummary workbook, target, batch
    stage = "WriteStrengthSummary"
    WriteStrengthSummary workbook, batch, units, section
    stage = "WriteCrackSummary"
    WriteCrackSummary workbook, batch, units
    stage = "WriteStabilitySummary"
    WriteStabilitySummary workbook, batch, units
    Exit Sub

Failed:
    Err.Raise Err.Number, "CBatchResultWriter.WriteSummary", stage & ": " & Err.Description
End Sub

' Очищает текущую сводку перед новым расчетом или ручной очисткой Results.
Public Sub ClearSummary(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3972, "CBatchResultWriter", "Книга Excel не передана."
    ClearSummaryBody SummaryClearRange(workbook)
    ClearStrengthSummary workbook
    ClearCrackSummary workbook
    ClearStabilitySummary workbook
End Sub

' Обновляет итоговое время уже после завершения записи Results и схемы.
' Сводка формируется перед детальным Results, потому что Plot.LoadCase = Worst читает из
' нее определяющее сочетание; поэтому общий elapsed time дозаписывается в конце.
Public Sub UpdateElapsedSeconds(ByVal workbook As Object, ByVal elapsedSeconds As Double)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3974, "CBatchResultWriter", "Книга Excel не передана."

    Dim target As Object
    Set target = SummaryTargetRange(workbook)
    target.Offset(3, 0).Resize(1, 1).Value2 = elapsedSeconds
End Sub

Private Function SummaryTargetRange(ByVal workbook As Object, Optional ByVal requiredRows As Long = 0) As Object
    Dim target As Object
    Set target = workbook.Names.Item("rngBatchSummary").RefersToRange(1, 1)
    If requiredRows <= 0 Then requiredRows = RequiredSummaryRows(1)
    If target.Row + requiredRows - 1 > target.Worksheet.Rows.Count Or _
            target.Column + TARGET_COLUMNS - 1 > target.Worksheet.Columns.Count Then
        Err.Raise vbObjectError + 3973, "CBatchResultWriter", _
            "rngBatchSummary не имеет места для требуемого размера сводки."
    End If
    Set SummaryTargetRange = target.Resize(requiredRows, TARGET_COLUMNS)
End Function

' Возвращает область безопасной очистки Summary до шапки следующего блока.
' Так старые строки от большего числа LC не остаются после уменьшения rngLoadCombinations.
Private Function SummaryClearRange(ByVal workbook As Object) As Object
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngBatchSummary").RefersToRange(1, 1)

    Dim rowsToClear As Long
    rowsToClear = RequiredSummaryOutputRows(1)

    On Error GoTo UseDefaultHeight
    Dim strengthWriter As CStrengthSummaryWriter
    Set strengthWriter = New CStrengthSummaryWriter
    Dim strengthAnchor As Object
    Set strengthAnchor = workbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange
    Dim strengthTopRow As Long
    strengthTopRow = strengthAnchor.Row - strengthWriter.HeaderRowsAboveAnchor
    If strengthTopRow > anchor.Row Then rowsToClear = MaxLong(rowsToClear, strengthTopRow - anchor.Row)

UseDefaultHeight:
    If anchor.Row + rowsToClear - 1 > anchor.Worksheet.Rows.Count Then rowsToClear = anchor.Worksheet.Rows.Count - anchor.Row + 1
    Set SummaryClearRange = anchor.Resize(rowsToClear, MaxLong(RequiredSummaryOutputColumns(), LEGACY_SUMMARY_CLEAR_COLUMNS))
End Function

' Снимает объединения и очищает текущую сводку до блочной записи массива.
' Это важно при повторном расчете: предыдущий запуск объединяет строки
' заголовков, и Excel не позволяет записывать двумерный массив поверх
' частично объединенного диапазона.
Private Sub PrepareSummaryTarget(ByVal target As Object)
    ClearSummaryBody target
End Sub

Private Sub ClearSummaryBody(ByVal target As Object)
    target.UnMerge
    target.Clear
End Sub

Private Sub WriteSummaryBody(ByVal target As Object, ByRef data() As Variant)
    target.Resize(UBound(data, 1), TARGET_COLUMNS).Value2 = data
End Sub

' Обновляет отдельную детальную таблицу устойчивости рядом с расчетным snapshot.
Private Sub WriteStabilitySummary(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    Dim writer As CStabilitySummaryWriter
    Set writer = New CStabilitySummaryWriter
    writer.WriteSummary workbook, batch, units
End Sub

' Обновляет отдельную детальную таблицу НДС и capacity по модели прочности.
Private Sub WriteStrengthSummary(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, _
        ByVal units As CUnitSystem, ByVal section As CSectionModel)
    Dim writer As CStrengthSummaryWriter
    Set writer = New CStrengthSummaryWriter
    writer.WriteSummary workbook, batch, units, section
End Sub

' Обновляет отдельную детальную таблицу нормальных и продольных трещин.
Private Sub WriteCrackSummary(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    Dim writer As CCrackSummaryWriter
    Set writer = New CCrackSummaryWriter
    writer.WriteSummary workbook, batch, units
End Sub

' Очищает отдельный блок прочности при общей очистке Results.
Private Sub ClearStrengthSummary(ByVal workbook As Object)
    Dim writer As CStrengthSummaryWriter
    Set writer = New CStrengthSummaryWriter
    writer.ClearSummary workbook
End Sub

' Очищает отдельный блок устойчивости при общей очистке Results.
Private Sub ClearStabilitySummary(ByVal workbook As Object)
    Dim writer As CStabilitySummaryWriter
    Set writer = New CStabilitySummaryWriter
    writer.ClearSummary workbook
End Sub

' Очищает отдельный блок трещин при общей очистке Results.
Private Sub ClearCrackSummary(ByVal workbook As Object)
    Dim writer As CCrackSummaryWriter
    Set writer = New CCrackSummaryWriter
    writer.ClearSummary workbook
End Sub

Private Sub PutMetaRows(ByRef data() As Variant, ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    data(1, 1) = batch.WorstCombinationID
    data(1, 2) = "Определяющее сочетание"
    data(2, 1) = batch.WorstCombinationName
    data(2, 2) = "Комментарий определяющего сочетания"
    data(3, 1) = batch.SolverCallCount
    data(3, 2) = "Количество решений НДС"
    data(4, 1) = batch.ElapsedSeconds
    data(4, 2) = "Время расчета, с"
End Sub

Private Sub PutGroupHeaders(ByRef data() As Variant)
    data(GROUP_HEADER_ROW, SUM_COL_STATUS_STRAIN) = "статус проверки"
    data(GROUP_HEADER_ROW, SUM_COL_RESERVE_STRAIN) = "минимальные коэффициенты запаса"
End Sub

Private Sub PutSubgroupHeaders(ByRef data() As Variant)
    data(SUBGROUP_HEADER_ROW, SUM_COL_STATUS_STRAIN) = "прочность"
    data(SUBGROUP_HEADER_ROW, SUM_COL_STATUS_CRACK) = "трещины"
    data(SUBGROUP_HEADER_ROW, SUM_COL_STATUS_STABILITY) = "устойчивость"
    data(SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_STRAIN) = "прочность"
    data(SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_CRACK) = "трещины"
    data(SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_STABILITY) = "устойчивость"
End Sub

Private Sub PutColumnHeaders(ByRef data() As Variant, ByVal units As CUnitSystem)
    data(COMMENT_HEADER_ROW, SUM_COL_ID) = "номер сочетания из таблицы нагрузок"
    data(COMMENT_HEADER_ROW, SUM_COL_COMMENT) = "пользовательский комментарий сочетания"
    data(COMMENT_HEADER_ROW, SUM_COL_OVERALL_STATUS) = "итоговый статус по всем включенным проверкам"
    data(COMMENT_HEADER_ROW, SUM_COL_STATUS_STRAIN) = "статус прямого НДС по модели прочности"
    data(COMMENT_HEADER_ROW, SUM_COL_STATUS_CAPACITY) = "статус поиска предельного состояния"
    data(COMMENT_HEADER_ROW, SUM_COL_STATUS_CRACK) = "статус нормальных трещин"
    data(COMMENT_HEADER_ROW, SUM_COL_STATUS_LONG_CRACK) = "статус продольных трещин"
    data(COMMENT_HEADER_ROW, SUM_COL_STATUS_STABILITY) = "общий статус устойчивости"
    data(COMMENT_HEADER_ROW, SUM_COL_RESERVE_STRAIN) = "минимальный запас по предельным деформациям"
    data(COMMENT_HEADER_ROW, SUM_COL_RESERVE_CAPACITY) = "запас по найденному множителю несущей"
    data(COMMENT_HEADER_ROW, SUM_COL_RESERVE_CRACK) = "a_crc,ult/a_crc"
    data(COMMENT_HEADER_ROW, SUM_COL_RESERVE_LONG_CRACK) = "Rb,mc2/sigma_b"
    data(COMMENT_HEADER_ROW, SUM_COL_RESERVE_STABILITY) = "минимальный запас по применимым плоскостям"

    data(COLUMN_HEADER_ROW, SUM_COL_ID) = "Combination ID"
    data(COLUMN_HEADER_ROW, SUM_COL_COMMENT) = "Comment"
    data(COLUMN_HEADER_ROW, SUM_COL_OVERALL_STATUS) = "общий статус"
    data(COLUMN_HEADER_ROW, SUM_COL_STATUS_STRAIN) = "по деформациям " & ChrW$(&H3B5)
    data(COLUMN_HEADER_ROW, SUM_COL_STATUS_CAPACITY) = "по предельным усилиям"
    data(COLUMN_HEADER_ROW, SUM_COL_STATUS_CRACK) = "нормальные трещины"
    data(COLUMN_HEADER_ROW, SUM_COL_STATUS_LONG_CRACK) = "продольные трещины"
    data(COLUMN_HEADER_ROW, SUM_COL_STATUS_STABILITY) = "устойчивость"
    data(COLUMN_HEADER_ROW, SUM_COL_RESERVE_STRAIN) = "по деформациям " & ChrW$(&H3B5)
    data(COLUMN_HEADER_ROW, SUM_COL_RESERVE_CAPACITY) = "по предельным усилиям"
    data(COLUMN_HEADER_ROW, SUM_COL_RESERVE_CRACK) = "нормальные трещины"
    data(COLUMN_HEADER_ROW, SUM_COL_RESERVE_LONG_CRACK) = "продольные трещины"
    data(COLUMN_HEADER_ROW, SUM_COL_RESERVE_STABILITY) = "устойчивость"
End Sub

Private Sub PutCombination(ByRef data() As Variant, ByVal rowIndex As Long, ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal units As CUnitSystem, ByVal section As CSectionModel)
    Dim policy As CBatchStatusPolicy
    Set policy = StatusPolicy()

    If rowIndex > UBound(data, 1) Then Exit Sub

    data(rowIndex, SUM_COL_ID) = batch.CombinationID(index)
    data(rowIndex, SUM_COL_COMMENT) = batch.CombinationName(index)
    data(rowIndex, SUM_COL_OVERALL_STATUS) = policy.DisplayStatus(batch.OverallStatus(index))

    PutStatusCell data, rowIndex, SUM_COL_STATUS_STRAIN, policy, _
        batch.StrengthDirectStateRequested(index), batch.DirectStateStatus(index)
    PutStatusCell data, rowIndex, SUM_COL_STATUS_CAPACITY, policy, _
        batch.StrengthCapacityRequested(index), batch.CapacityStatus(index)
    PutStatusCell data, rowIndex, SUM_COL_STATUS_CRACK, policy, _
        batch.CrackWidthRequested(index), batch.CrackStatus(index)
    PutStatusCell data, rowIndex, SUM_COL_STATUS_LONG_CRACK, policy, _
        batch.CrackWidthRequested(index), batch.LongitudinalCrackStatus(index)
    PutStatusCell data, rowIndex, SUM_COL_STATUS_STABILITY, policy, _
        batch.StabilityRequested(index), batch.StabilityStatus(index)

    PutReserveCell data, rowIndex, SUM_COL_RESERVE_STRAIN, batch.StrengthStrainReserve(index)
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_CAPACITY, batch.StrengthCapacityReserve(index)
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_CRACK, batch.CrackReserveFactor(index)
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_LONG_CRACK, batch.LongitudinalCrackReserveFactor(index)
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_STABILITY, batch.StabilityReserveForSummary(index)
End Sub

' Записывает статус проверки или N/A, если расчетная ветка не была включена профилем.
Private Sub PutStatusCell(ByRef data() As Variant, ByVal rowIndex As Long, ByVal columnIndex As Long, _
        ByVal policy As CBatchStatusPolicy, ByVal requested As Boolean, ByVal statusText As String)
    If requested Then
        data(rowIndex, columnIndex) = policy.DisplayStatus(statusText)
    Else
        data(rowIndex, columnIndex) = policy.NotApplicable
    End If
End Sub

' Записывает только положительные расчетные коэффициенты запаса.
Private Sub PutReserveCell(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal columnIndex As Long, ByVal reserveValue As Double)
    If reserveValue > 0# Then data(rowIndex, columnIndex) = reserveValue
End Sub

Private Sub FormatSummary(ByVal workbook As Object, ByVal target As Object, ByVal batch As CBatchSectionCalculator)
    On Error GoTo Failed

    Dim stage As String
    stage = "ResetRange"
    target.UnMerge
    target.Interior.ColorIndex = -4142
    target.Borders.LineStyle = -4142

    stage = "BaseFont"
    target.HorizontalAlignment = -4131
    target.VerticalAlignment = -4108
    target.WrapText = False
    target.Font.Name = "Arial"
    target.Font.Size = 9
    target.Rows.RowHeight = 15
    target.Worksheet.Columns.ColumnWidth = RESULTS_COLUMN_WIDTH

    stage = "HeaderStyle"
    target.Offset(0, 0).Resize(META_ROWS, 5).Interior.Color = HEADER_COLOR
    target.Offset(0, 1).Resize(META_ROWS, 4).Font.Italic = True
    target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(COMMENT_HEADER_ROW - GROUP_HEADER_ROW + 1, TARGET_COLUMNS).Interior.Color = HEADER_COLOR
    target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(COMMENT_HEADER_ROW - GROUP_HEADER_ROW, TARGET_COLUMNS).Font.Bold = True
    target.Rows.Item(COMMENT_HEADER_ROW).Font.Bold = False

    stage = "MergeHeaders"
    MergeSummaryBlock target, 1, 1, 2, 5
    MergeSummaryBlock target, 2, 2, 2, 5
    MergeSummaryBlock target, 3, 3, 2, 5
    MergeSummaryBlock target, 4, 4, 2, 5
    MergeSummaryBlock target, GROUP_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_ID, SUM_COL_ID
    MergeSummaryBlock target, GROUP_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_COMMENT, SUM_COL_COMMENT
    MergeSummaryBlock target, GROUP_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_OVERALL_STATUS, SUM_COL_OVERALL_STATUS
    MergeSummaryHeader target, GROUP_HEADER_ROW, SUM_COL_STATUS_STRAIN, SUM_COL_STATUS_STABILITY
    MergeSummaryHeader target, GROUP_HEADER_ROW, SUM_COL_RESERVE_STRAIN, SUM_COL_RESERVE_STABILITY
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, SUM_COL_STATUS_STRAIN, SUM_COL_STATUS_CAPACITY
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, SUM_COL_STATUS_CRACK, SUM_COL_STATUS_LONG_CRACK
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_STRAIN, SUM_COL_RESERVE_CAPACITY
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_CRACK, SUM_COL_RESERVE_LONG_CRACK
    MergeSummaryBlock target, SUBGROUP_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_STATUS_STABILITY, SUM_COL_STATUS_STABILITY
    MergeSummaryBlock target, SUBGROUP_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_RESERVE_STABILITY, SUM_COL_RESERVE_STABILITY
    MergeColumnLabels target, COLUMN_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_STATUS_STRAIN, SUM_COL_STATUS_CAPACITY
    MergeColumnLabels target, COLUMN_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_STATUS_CRACK, SUM_COL_STATUS_LONG_CRACK
    MergeColumnLabels target, COLUMN_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_RESERVE_STRAIN, SUM_COL_RESERVE_CAPACITY
    MergeColumnLabels target, COLUMN_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_RESERVE_CRACK, SUM_COL_RESERVE_LONG_CRACK

    stage = "Alignment"
    target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(COMMENT_HEADER_ROW - GROUP_HEADER_ROW + 1, TARGET_COLUMNS).HorizontalAlignment = -4108
    target.Rows.Item(COMMENT_HEADER_ROW).HorizontalAlignment = -4131
    target.Offset(0, 0).Resize(META_ROWS, 1).HorizontalAlignment = -4108
    target.Offset(0, 1).Resize(META_ROWS, 4).HorizontalAlignment = -4131
    target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(COMMENT_HEADER_ROW - GROUP_HEADER_ROW, TARGET_COLUMNS).WrapText = True
    target.Rows.Item(COMMENT_HEADER_ROW).WrapText = False

    Dim dataRows As Long
    dataRows = target.Rows.Count - FIRST_DATA_ROW + 1
    If dataRows < 1 Then dataRows = 1
    target.Offset(FIRST_DATA_ROW - 1, 0).Resize(dataRows, TARGET_COLUMNS).HorizontalAlignment = -4108
    target.Offset(FIRST_DATA_ROW - 1, SUM_COL_COMMENT - 1).Resize(dataRows, 1).HorizontalAlignment = -4131
    target.Offset(FIRST_DATA_ROW - 1, SUM_COL_RESERVE_STRAIN - 1).Resize(dataRows, _
        SUM_COL_RESERVE_STABILITY - SUM_COL_RESERVE_STRAIN + 1).Interior.Color = SAFETY_COLOR
    target.Offset(GROUP_HEADER_ROW - 1, 3).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, 1).Interior.ColorIndex = -4142
    target.Offset(GROUP_HEADER_ROW - 1, 9).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, 1).Interior.ColorIndex = -4142
    stage = "FormatOverallStatusRows"
    FormatOverallStatusRows target, batch

    stage = "Borders"
    ApplySummaryBorders target.Offset(0, 0).Resize(META_ROWS, 5)
    ApplySummaryBorders target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, 3)
    ApplySummaryBorders target.Offset(GROUP_HEADER_ROW - 1, SUM_COL_STATUS_STRAIN - 1).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, 5)
    ApplySummaryBorders target.Offset(GROUP_HEADER_ROW - 1, SUM_COL_RESERVE_STRAIN - 1).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, 5)
    stage = "NumberFormat"
    If dataRows > 0 Then
        target.Offset(FIRST_DATA_ROW - 1, SUM_COL_RESERVE_STRAIN - 1).Resize(dataRows, _
            SUM_COL_RESERVE_STABILITY - SUM_COL_RESERVE_STRAIN + 1).NumberFormat = "0.000"
    End If
    Exit Sub

Failed:
    Err.Raise Err.Number, "CBatchResultWriter.FormatSummary", stage & ": " & Err.Description
End Sub

Private Sub MergeSummaryHeader(ByVal target As Object, ByVal rowIndex As Long, ByVal firstColumn As Long, ByVal lastColumn As Long)
    Dim area As Object
    Set area = target.Offset(rowIndex - 1, firstColumn - 1).Resize(1, lastColumn - firstColumn + 1)
    area.Merge
End Sub

Private Sub MergeSummaryBlock(ByVal target As Object, ByVal firstRow As Long, ByVal lastRow As Long, _
        ByVal firstColumn As Long, ByVal lastColumn As Long)
    Dim area As Object
    Set area = target.Offset(firstRow - 1, firstColumn - 1).Resize(lastRow - firstRow + 1, lastColumn - firstColumn + 1)
    area.Merge
End Sub

' Объединяет вертикальные ячейки обозначений в шапке для группы соседних столбцов.
Private Sub MergeColumnLabels(ByVal target As Object, ByVal firstRow As Long, ByVal lastRow As Long, _
        ByVal firstColumn As Long, ByVal lastColumn As Long)
    Dim columnIndex As Long
    For columnIndex = firstColumn To lastColumn
        MergeSummaryBlock target, firstRow, lastRow, columnIndex, columnIndex
    Next columnIndex
End Sub

Private Sub ApplySummaryBorders(ByVal area As Object)
    area.Borders.LineStyle = 1
    area.Borders.Color = RGB(0, 0, 0)
    area.Borders.Weight = 2
End Sub

' Подсвечивает только конкретные ячейки проверок, которые не прошли.
' Соседние блоки LC остаются без красной заливки, если их собственный статус OK.
Private Sub FormatOverallStatusRows(ByVal target As Object, ByVal batch As CBatchSectionCalculator)
    If batch Is Nothing Then Exit Sub
    If batch.Count <= 0 Then Exit Sub

    Dim policy As CBatchStatusPolicy
    Set policy = StatusPolicy()

    Dim i As Long
    For i = 1 To batch.Count
        Dim rowNumber As Long
        rowNumber = FIRST_DATA_ROW + i - 2
        HighlightFailedCell target, rowNumber, SUM_COL_OVERALL_STATUS, policy, batch.OverallStatus(i)
        HighlightFailedPair target, rowNumber, SUM_COL_STATUS_STRAIN, SUM_COL_RESERVE_STRAIN, policy, batch.DirectStateStatus(i)
        HighlightFailedPair target, rowNumber, SUM_COL_STATUS_CAPACITY, SUM_COL_RESERVE_CAPACITY, policy, batch.CapacityStatus(i)
        HighlightFailedPair target, rowNumber, SUM_COL_STATUS_CRACK, SUM_COL_RESERVE_CRACK, policy, batch.CrackStatus(i)
        HighlightFailedPair target, rowNumber, SUM_COL_STATUS_LONG_CRACK, SUM_COL_RESERVE_LONG_CRACK, policy, batch.LongitudinalCrackStatus(i)
        HighlightFailedPair target, rowNumber, SUM_COL_STATUS_STABILITY, SUM_COL_RESERVE_STABILITY, policy, batch.StabilityStatus(i)
    Next i
End Sub

Private Sub HighlightFailedPair(ByVal target As Object, ByVal rowNumber As Long, _
        ByVal statusColumn As Long, ByVal reserveColumn As Long, _
        ByVal policy As CBatchStatusPolicy, ByVal statusText As String)
    If Not policy.IsFailedOverall(statusText) Then Exit Sub
    HighlightCell target, rowNumber, statusColumn
    HighlightCell target, rowNumber, reserveColumn
End Sub

Private Sub HighlightFailedCell(ByVal target As Object, ByVal rowNumber As Long, _
        ByVal columnIndex As Long, ByVal policy As CBatchStatusPolicy, ByVal statusText As String)
    If policy.IsFailedOverall(statusText) Then HighlightCell target, rowNumber, columnIndex
End Sub

Private Sub HighlightCell(ByVal target As Object, ByVal rowNumber As Long, ByVal columnIndex As Long)
    With target.Offset(rowNumber, columnIndex - 1).Resize(1, 1)
        .Interior.Color = FAILED_ROW_COLOR
        .Font.Color = FAILED_FONT_COLOR
    End With
End Sub

Private Function StatusPolicy() As CBatchStatusPolicy
    Static policy As CBatchStatusPolicy
    If policy Is Nothing Then Set policy = New CBatchStatusPolicy
    Set StatusPolicy = policy
End Function

Private Function MaxLong(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxLong = a Else MaxLong = b
End Function
