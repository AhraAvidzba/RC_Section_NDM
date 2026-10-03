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

Private Const TARGET_COLUMNS As Long = 26 ' Ширина верхней сводки с разложением устойчивости по СП, плоскостям и ветвям.
Private Const META_ROWS As Long = 4 ' Верхние строки с общими сведениями о запуске.
Private Const GROUP_HEADER_ROW As Long = 6 ' Первая строка шестистрочной шапки основной таблицы.
Private Const SUBGROUP_HEADER_ROW As Long = 7 ' Строка внутренних подзаголовков расчетных блоков.
Private Const COLUMN_HEADER_ROW As Long = 8 ' Верх строки с обозначениями конкретных параметров.
Private Const COLUMN_HEADER_LAST_ROW As Long = 10 ' Низ объединенных ячеек с обозначениями параметров.
Private Const COMMENT_HEADER_ROW As Long = 11 ' Последняя строка шапки: короткие комментарии к столбцам.
Private Const WORST_HEADER_ROW As Long = 12 ' Строка с худшим LC и минимальным запасом по каждой проверке.
Private Const FIRST_DATA_ROW As Long = 13 ' Первая строка сочетаний.

Private Const SUM_COL_ID As Long = 1 ' CombinationID в новой компактной сводке.
Private Const SUM_COL_COMMENT As Long = 2 ' Пользовательский комментарий LC.
Private Const SUM_COL_RESULT_COMMENT As Long = 3 ' Пояснение итогового результата LC из result-subtree.
Private Const SUM_COL_OVERALL_STATUS As Long = 4 ' Итоговый статус LC по всем включенным проверкам.
Private Const SUM_COL_STATUS_STRAIN As Long = 6 ' Статус прямого НДС по модели прочности.
Private Const SUM_COL_STATUS_CAPACITY As Long = 7 ' Статус проверки несущей способности.
Private Const SUM_COL_STATUS_CRACK As Long = 8 ' Статус нормальных трещин.
Private Const SUM_COL_STATUS_LONG_CRACK As Long = 9 ' Статус продольных трещин.
Private Const SUM_COL_STATUS_SP35_P1_ETA As Long = 10 ' СП35, плоскость 1, ветвь ec > r.
Private Const SUM_COL_STATUS_SP35_P1_TABLE As Long = 11 ' СП35, плоскость 1, ветвь ec <= r.
Private Const SUM_COL_STATUS_SP35_P2_ETA As Long = 12 ' СП35, плоскость 2, ветвь ec > r.
Private Const SUM_COL_STATUS_SP35_P2_TABLE As Long = 13 ' СП35, плоскость 2, ветвь ec <= r.
Private Const SUM_COL_STATUS_SP63_P1 As Long = 14 ' СП63, плоскость 1.
Private Const SUM_COL_STATUS_SP63_P2 As Long = 15 ' СП63, плоскость 2.
Private Const SUM_COL_RESERVE_STRAIN As Long = 17 ' Запас по деформациям strength-модели.
Private Const SUM_COL_RESERVE_CAPACITY As Long = 18 ' Запас по предельным усилиям.
Private Const SUM_COL_RESERVE_CRACK As Long = 19 ' Запас по нормальным трещинам.
Private Const SUM_COL_RESERVE_LONG_CRACK As Long = 20 ' Запас по продольным трещинам.
Private Const SUM_COL_RESERVE_SP35_P1_ETA As Long = 21 ' Запас СП35, плоскость 1, ec > r: limit*Ncr/N.
Private Const SUM_COL_RESERVE_SP35_P1_TABLE As Long = 22 ' Запас СП35, плоскость 1, ec <= r: Nult/N.
Private Const SUM_COL_RESERVE_SP35_P2_ETA As Long = 23 ' Запас СП35, плоскость 2, ec > r: limit*Ncr/N.
Private Const SUM_COL_RESERVE_SP35_P2_TABLE As Long = 24 ' Запас СП35, плоскость 2, ec <= r: Nult/N.
Private Const SUM_COL_RESERVE_SP63_P1 As Long = 25 ' Запас СП63, плоскость 1: Ncr/N.
Private Const SUM_COL_RESERVE_SP63_P2 As Long = 26 ' Запас СП63, плоскость 2: Ncr/N.

Private Const STATUS_LEGEND_GAP_COLUMNS As Long = 1 ' Один пустой столбец между summary и словарем статусов.
Private Const STATUS_LEGEND_COLUMNS As Long = 2 ' Ширина правого словаря статусов: код и текстовая расшифровка.
Private Const STATUS_LEGEND_ROWS As Long = 9 ' Заголовок, шапка и семь пользовательских статусов.

Private Const DEFAULT_SUMMARY_ROWS As Long = 42 ' Штатная высота шаблона: 12 строк до данных и 30 строк LC.

Private Const HEADER_COLOR As Long = &HD9D9D9 ' Светло-серый фон заголовков.

' Возвращает высоту основной таблицы Summary для фактического количества LC.
Public Function RequiredSummaryRows(ByVal combinationCount As Long) As Long
    If combinationCount < 0 Then combinationCount = 0
    RequiredSummaryRows = FIRST_DATA_ROW - 1 + combinationCount
    If RequiredSummaryRows < COLUMN_HEADER_ROW Then RequiredSummaryRows = COLUMN_HEADER_ROW
    If RequiredSummaryRows < DEFAULT_SUMMARY_ROWS Then RequiredSummaryRows = DEFAULT_SUMMARY_ROWS
End Function

' Возвращает число строк LC по высоте rngLoadCombinations. Это нужно для
' пустой сетки summary: пользователь видит все доступные строки сочетаний,
' даже если часть из них сейчас не заполнена.
Public Function RequiredDataRowsForWorkbook(ByVal workbook As Object, ByVal combinationCount As Long) As Long
    RequiredDataRowsForWorkbook = combinationCount
    If RequiredDataRowsForWorkbook < 1 Then RequiredDataRowsForWorkbook = 1

    On Error GoTo SafeExit
    If workbook Is Nothing Then GoTo SafeExit
    Dim loads As Object
    Set loads = workbook.Names.Item("rngLoadCombinations").RefersToRange
    If loads.Rows.Count > 1 Then
        If loads.Rows.Count - 1 > RequiredDataRowsForWorkbook Then _
            RequiredDataRowsForWorkbook = loads.Rows.Count - 1
    End If
SafeExit:
    If RequiredDataRowsForWorkbook < 1 Then RequiredDataRowsForWorkbook = 1
End Function

' Возвращает высоту summary для фактической книги с учетом доступных строк rngLoadCombinations.
Public Function RequiredSummaryRowsForWorkbook(ByVal workbook As Object, ByVal combinationCount As Long) As Long
    RequiredSummaryRowsForWorkbook = RequiredSummaryRows(RequiredDataRowsForWorkbook(workbook, combinationCount))
End Function

' Возвращает ширину собственно rngBatchSummary.
Public Function RequiredSummaryColumns() As Long
    RequiredSummaryColumns = TARGET_COLUMNS
End Function

' Возвращает полную ширину, которую writer затрагивает на Results.
Public Function RequiredSummaryOutputColumns() As Long
    RequiredSummaryOutputColumns = TARGET_COLUMNS + STATUS_LEGEND_GAP_COLUMNS + STATUS_LEGEND_COLUMNS
End Function

' Возвращает полную высоту, которую занимает компактная сводка.
Public Function RequiredSummaryOutputRows(ByVal combinationCount As Long) As Long
    RequiredSummaryOutputRows = RequiredSummaryRows(combinationCount)
End Function

' Возвращает полную высоту summary для проверки раскладки Results в фактической книге.
Public Function RequiredSummaryOutputRowsForWorkbook(ByVal workbook As Object, ByVal combinationCount As Long) As Long
    RequiredSummaryOutputRowsForWorkbook = RequiredSummaryRowsForWorkbook(workbook, combinationCount)
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
    summaryRows = RequiredSummaryRowsForWorkbook(workbook, batch.Count)

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
    PutColumnHeaders data, units, batch
    stage = "PutWorstRow"
    PutWorstRow data, batch

    Dim i As Long
    For i = 1 To batch.Count
        stage = "PutCombination " & CStr(i)
        PutCombination data, FIRST_DATA_ROW + OutputDataOffset(batch, i, summaryRows - FIRST_DATA_ROW + 1) - 1, batch, i, units, section
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

' Возвращает область rngBatchSummary нужного размера и сразу проверяет, хватает ли листа Results.
Private Function SummaryTargetRange(ByVal workbook As Object, Optional ByVal requiredRows As Long = 0) As Object
    Dim target As Object
    Set target = workbook.Names.Item("rngBatchSummary").RefersToRange(1, 1)
    If requiredRows <= 0 Then requiredRows = RequiredSummaryRows(1)
    If target.Row + requiredRows - 1 > target.Worksheet.Rows.Count Or _
            target.Column + RequiredSummaryOutputColumns() - 1 > target.Worksheet.Columns.Count Then
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
    Set SummaryClearRange = anchor.Resize(rowsToClear, RequiredSummaryOutputColumns())
End Function

' Снимает объединения и очищает текущую сводку до блочной записи массива.
' Это важно при повторном расчете: предыдущий запуск объединяет строки
' заголовков, и Excel не позволяет записывать двумерный массив поверх
' частично объединенного диапазона.
Private Sub PrepareSummaryTarget(ByVal target As Object)
    ClearSummaryBody target
End Sub

' Полностью очищает область summary, включая старые объединения и форматы.
Private Sub ClearSummaryBody(ByVal target As Object)
    target.UnMerge
    target.Clear
End Sub

' Записывает уже подготовленный двумерный массив в Excel одним присваиванием.
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

' Заполняет четыре верхние строки summary: worst-сочетание, комментарий,
' число вызовов CSectionSolver и общее время расчета.
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

' Заполняет верхний уровень шапки compact summary: статусы и коэффициенты
' запаса как два независимых смысловых блока.
Private Sub PutGroupHeaders(ByRef data() As Variant)
    data(GROUP_HEADER_ROW, SUM_COL_STATUS_STRAIN) = "статус проверки"
    data(GROUP_HEADER_ROW, SUM_COL_RESERVE_STRAIN) = "минимальные коэффициенты запаса"
End Sub

' Заполняет второй уровень шапки: прочность, трещины и устойчивость.
Private Sub PutSubgroupHeaders(ByRef data() As Variant)
    data(SUBGROUP_HEADER_ROW, SUM_COL_STATUS_STRAIN) = "прочность"
    data(SUBGROUP_HEADER_ROW, SUM_COL_STATUS_CRACK) = "трещины"
    data(SUBGROUP_HEADER_ROW, SUM_COL_STATUS_SP35_P1_ETA) = "устойчивость по СП35"
    data(SUBGROUP_HEADER_ROW, SUM_COL_STATUS_SP63_P1) = "устойчивость по СП63"
    data(SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_STRAIN) = "прочность"
    data(SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_CRACK) = "трещины"
    data(SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_SP35_P1_ETA) = "устойчивость по СП35"
    data(SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_SP63_P1) = "устойчивость по СП63"
End Sub

' Заполняет строки обозначений и комментариев к каждому столбцу summary.
Private Sub PutColumnHeaders(ByRef data() As Variant, ByVal units As CUnitSystem, ByVal batch As CBatchSectionCalculator)
    Dim sp35EtaLabel As String
    sp35EtaLabel = SP35EtaReserveLabel(batch)

    data(COMMENT_HEADER_ROW, SUM_COL_ID) = "номер сочетания из таблицы нагрузок"
    data(COMMENT_HEADER_ROW, SUM_COL_COMMENT) = "пользовательский комментарий сочетания"
    data(COMMENT_HEADER_ROW, SUM_COL_RESULT_COMMENT) = "пояснение итогового расчетного статуса"
    data(COMMENT_HEADER_ROW, SUM_COL_OVERALL_STATUS) = "итоговый статус по всем включенным проверкам"
    data(COMMENT_HEADER_ROW, SUM_COL_STATUS_STRAIN) = "статус прямого НДС по модели прочности"
    data(COMMENT_HEADER_ROW, SUM_COL_STATUS_CAPACITY) = "статус поиска предельного состояния"
    data(COMMENT_HEADER_ROW, SUM_COL_STATUS_CRACK) = "статус нормальных трещин"
    data(COMMENT_HEADER_ROW, SUM_COL_STATUS_LONG_CRACK) = "статус продольных трещин"
    data(COMMENT_HEADER_ROW, SUM_COL_STATUS_SP35_P1_ETA) = "СП35: плоскость 1, ветвь ec > r"
    data(COMMENT_HEADER_ROW, SUM_COL_STATUS_SP35_P1_TABLE) = "СП35: плоскость 1, ветвь ec <= r"
    data(COMMENT_HEADER_ROW, SUM_COL_STATUS_SP35_P2_ETA) = "СП35: плоскость 2, ветвь ec > r"
    data(COMMENT_HEADER_ROW, SUM_COL_STATUS_SP35_P2_TABLE) = "СП35: плоскость 2, ветвь ec <= r"
    data(COMMENT_HEADER_ROW, SUM_COL_STATUS_SP63_P1) = "СП63: плоскость 1"
    data(COMMENT_HEADER_ROW, SUM_COL_STATUS_SP63_P2) = "СП63: плоскость 2"
    data(COMMENT_HEADER_ROW, SUM_COL_RESERVE_STRAIN) = "минимальный запас по предельным деформациям"
    data(COMMENT_HEADER_ROW, SUM_COL_RESERVE_CAPACITY) = "запас по найденному множителю несущей"
    data(COMMENT_HEADER_ROW, SUM_COL_RESERVE_CRACK) = "a_crc,ult/a_crc"
    data(COMMENT_HEADER_ROW, SUM_COL_RESERVE_LONG_CRACK) = "Rb,mc2/sigma_b"
    data(COMMENT_HEADER_ROW, SUM_COL_RESERVE_SP35_P1_ETA) = "СП35: плоскость 1, запас " & sp35EtaLabel
    data(COMMENT_HEADER_ROW, SUM_COL_RESERVE_SP35_P1_TABLE) = "СП35: плоскость 1, запас Nult/N"
    data(COMMENT_HEADER_ROW, SUM_COL_RESERVE_SP35_P2_ETA) = "СП35: плоскость 2, запас " & sp35EtaLabel
    data(COMMENT_HEADER_ROW, SUM_COL_RESERVE_SP35_P2_TABLE) = "СП35: плоскость 2, запас Nult/N"
    data(COMMENT_HEADER_ROW, SUM_COL_RESERVE_SP63_P1) = "СП63: плоскость 1, запас Ncr/N"
    data(COMMENT_HEADER_ROW, SUM_COL_RESERVE_SP63_P2) = "СП63: плоскость 2, запас Ncr/N"

    data(COLUMN_HEADER_ROW, SUM_COL_ID) = "Combination ID"
    data(COLUMN_HEADER_ROW, SUM_COL_COMMENT) = "Comment"
    data(COLUMN_HEADER_ROW, SUM_COL_RESULT_COMMENT) = "ResultComment"
    data(COLUMN_HEADER_ROW, SUM_COL_OVERALL_STATUS) = "общий статус"
    data(COLUMN_HEADER_ROW, SUM_COL_STATUS_STRAIN) = "по деформациям " & ChrW$(&H3B5)
    data(COLUMN_HEADER_ROW, SUM_COL_STATUS_CAPACITY) = "по предельным усилиям"
    data(COLUMN_HEADER_ROW, SUM_COL_STATUS_CRACK) = "нормальные трещины"
    data(COLUMN_HEADER_ROW, SUM_COL_STATUS_LONG_CRACK) = "продольные трещины"
    PutStabilityColumnHeaders data, SUM_COL_STATUS_SP35_P1_ETA, SUM_COL_STATUS_SP63_P1, sp35EtaLabel
    data(COLUMN_HEADER_ROW, SUM_COL_RESERVE_STRAIN) = "по деформациям " & ChrW$(&H3B5)
    data(COLUMN_HEADER_ROW, SUM_COL_RESERVE_CAPACITY) = "по предельным усилиям"
    data(COLUMN_HEADER_ROW, SUM_COL_RESERVE_CRACK) = "нормальные трещины"
    data(COLUMN_HEADER_ROW, SUM_COL_RESERVE_LONG_CRACK) = "продольные трещины"
    PutStabilityColumnHeaders data, SUM_COL_RESERVE_SP35_P1_ETA, SUM_COL_RESERVE_SP63_P1, sp35EtaLabel
End Sub

' Заполняет трехстрочную часть шапки для устойчивости: плоскость, ветвь и
' проверяемый коэффициент запаса. Для СП35 eta-ветви показывается именно
' пользовательский лимит N/Ncr, а не условное обозначение k.
Private Sub PutStabilityColumnHeaders(ByRef data() As Variant, ByVal sp35FirstColumn As Long, _
        ByVal sp63FirstColumn As Long, ByVal sp35EtaLabel As String)
    data(COLUMN_HEADER_ROW, sp35FirstColumn) = "Плоскость 1"
    data(COLUMN_HEADER_ROW, sp35FirstColumn + 2) = "Плоскость 2"
    data(COLUMN_HEADER_ROW + 1, sp35FirstColumn) = "ec>r"
    data(COLUMN_HEADER_ROW + 1, sp35FirstColumn + 1) = "ec<=r"
    data(COLUMN_HEADER_ROW + 1, sp35FirstColumn + 2) = "ec>r"
    data(COLUMN_HEADER_ROW + 1, sp35FirstColumn + 3) = "ec<=r"
    data(COLUMN_HEADER_ROW + 2, sp35FirstColumn) = sp35EtaLabel
    data(COLUMN_HEADER_ROW + 2, sp35FirstColumn + 1) = "Nult/N"
    data(COLUMN_HEADER_ROW + 2, sp35FirstColumn + 2) = sp35EtaLabel
    data(COLUMN_HEADER_ROW + 2, sp35FirstColumn + 3) = "Nult/N"

    data(COLUMN_HEADER_ROW, sp63FirstColumn) = "Плоскость 1"
    data(COLUMN_HEADER_ROW, sp63FirstColumn + 1) = "Плоскость 2"
    data(COLUMN_HEADER_ROW + 2, sp63FirstColumn) = "Ncr/N"
    data(COLUMN_HEADER_ROW + 2, sp63FirstColumn + 1) = "Ncr/N"
End Sub

' Заполняет строку "worst LC" отдельно по каждой проверке. В левом блоке
' выводится номер сочетания, где запас минимален, а в правом блоке - сам
' этот минимальный запас. Так сводка сразу показывает, какое LC управляет
' каждой конкретной проверкой.
Private Sub PutWorstRow(ByRef data() As Variant, ByVal batch As CBatchSectionCalculator)
    data(WORST_HEADER_ROW, SUM_COL_ID) = "worst LC"
    data(WORST_HEADER_ROW, SUM_COL_COMMENT) = "-"
    data(WORST_HEADER_ROW, SUM_COL_RESULT_COMMENT) = "-"
    data(WORST_HEADER_ROW, SUM_COL_OVERALL_STATUS) = "-"
    PutWorstCheck data, batch, SUM_COL_STATUS_STRAIN, SUM_COL_RESERVE_STRAIN, "StrengthStrain"
    PutWorstCheck data, batch, SUM_COL_STATUS_CAPACITY, SUM_COL_RESERVE_CAPACITY, "StrengthCapacity"
    PutWorstCheck data, batch, SUM_COL_STATUS_CRACK, SUM_COL_RESERVE_CRACK, "Crack"
    PutWorstCheck data, batch, SUM_COL_STATUS_LONG_CRACK, SUM_COL_RESERVE_LONG_CRACK, "LongCrack"
    PutWorstCheck data, batch, SUM_COL_STATUS_SP35_P1_ETA, SUM_COL_RESERVE_SP35_P1_ETA, "SP35P1Eta"
    PutWorstCheck data, batch, SUM_COL_STATUS_SP35_P1_TABLE, SUM_COL_RESERVE_SP35_P1_TABLE, "SP35P1Table"
    PutWorstCheck data, batch, SUM_COL_STATUS_SP35_P2_ETA, SUM_COL_RESERVE_SP35_P2_ETA, "SP35P2Eta"
    PutWorstCheck data, batch, SUM_COL_STATUS_SP35_P2_TABLE, SUM_COL_RESERVE_SP35_P2_TABLE, "SP35P2Table"
    PutWorstCheck data, batch, SUM_COL_STATUS_SP63_P1, SUM_COL_RESERVE_SP63_P1, "SP63P1"
    PutWorstCheck data, batch, SUM_COL_STATUS_SP63_P2, SUM_COL_RESERVE_SP63_P2, "SP63P2"
End Sub

' Выбирает худшее LC для одного столбца проверки по минимальному положительному
' запасу. Неприменимые ветви и пустые запасы не участвуют в выборе.
Private Sub PutWorstCheck(ByRef data() As Variant, ByVal batch As CBatchSectionCalculator, _
        ByVal statusColumn As Long, ByVal reserveColumn As Long, ByVal checkKey As String)
    Dim worstIndex As Long
    Dim worstReserve As Double
    Dim i As Long
    For i = 1 To batch.Count
        Dim reserveValue As Double
        reserveValue = SummaryReserveByKey(batch, i, checkKey)
        If reserveValue > 0# Then
            If worstIndex <= 0 Or reserveValue < worstReserve Then
                worstIndex = i
                worstReserve = reserveValue
            End If
        End If
    Next i

    If worstIndex > 0 Then
        data(WORST_HEADER_ROW, statusColumn) = batch.CombinationID(worstIndex)
        data(WORST_HEADER_ROW, reserveColumn) = worstReserve
    End If
End Sub

' Записывает одну строку сочетания в compact summary. Все статусы идут через
' CResultStatusPolicy, а коэффициенты берутся из подробных расчетных результатов
' batch, чтобы summary совпадала с детальными таблицами ниже.
Private Sub PutCombination(ByRef data() As Variant, ByVal rowIndex As Long, ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal units As CUnitSystem, ByVal section As CSectionModel)
    Dim policy As CResultStatusPolicy
    Set policy = StatusPolicy()

    If rowIndex > UBound(data, 1) Then Exit Sub

    data(rowIndex, SUM_COL_ID) = batch.CombinationID(index)
    data(rowIndex, SUM_COL_COMMENT) = batch.CombinationName(index)
    data(rowIndex, SUM_COL_RESULT_COMMENT) = MetaComment(batch.ResultAt(index).OverallMeta)
    data(rowIndex, SUM_COL_OVERALL_STATUS) = MetaStatus(policy, batch.ResultAt(index).OverallMeta)

    PutStatusCell data, rowIndex, SUM_COL_STATUS_STRAIN, policy, _
        batch.StrengthDirectStateRequested(index), batch.ResultAt(index).DirectStateMeta
    PutStatusCell data, rowIndex, SUM_COL_STATUS_CAPACITY, policy, _
        batch.StrengthCapacityRequested(index), batch.ResultAt(index).CapacityMeta
    PutStatusCell data, rowIndex, SUM_COL_STATUS_CRACK, policy, _
        batch.CrackWidthRequested(index), batch.ResultAt(index).CrackMeta
    PutStatusCell data, rowIndex, SUM_COL_STATUS_LONG_CRACK, policy, _
        batch.CrackWidthRequested(index), batch.ResultAt(index).LongitudinalCrackMeta
    PutStabilityBranchStatus data, rowIndex, SUM_COL_STATUS_SP35_P1_ETA, policy, batch, index, 1, "SP35-eta"
    PutStabilityBranchStatus data, rowIndex, SUM_COL_STATUS_SP35_P1_TABLE, policy, batch, index, 1, "SP35-table"
    PutStabilityBranchStatus data, rowIndex, SUM_COL_STATUS_SP35_P2_ETA, policy, batch, index, 2, "SP35-eta"
    PutStabilityBranchStatus data, rowIndex, SUM_COL_STATUS_SP35_P2_TABLE, policy, batch, index, 2, "SP35-table"
    PutStabilityBranchStatus data, rowIndex, SUM_COL_STATUS_SP63_P1, policy, batch, index, 1, "SP63"
    PutStabilityBranchStatus data, rowIndex, SUM_COL_STATUS_SP63_P2, policy, batch, index, 2, "SP63"

    PutReserveCell data, rowIndex, SUM_COL_RESERVE_STRAIN, batch.ResultAt(index).StrengthResult.DirectState.StrainReserve
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_CAPACITY, batch.ResultAt(index).StrengthResult.Capacity.ReserveFactor
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_CRACK, batch.ResultAt(index).CrackResult.Width.ReserveFactor
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_LONG_CRACK, batch.ResultAt(index).CrackResult.Longitudinal.ReserveFactor
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_SP35_P1_ETA, SummaryReserveByKey(batch, index, "SP35P1Eta")
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_SP35_P1_TABLE, SummaryReserveByKey(batch, index, "SP35P1Table")
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_SP35_P2_ETA, SummaryReserveByKey(batch, index, "SP35P2Eta")
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_SP35_P2_TABLE, SummaryReserveByKey(batch, index, "SP35P2Table")
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_SP63_P1, SummaryReserveByKey(batch, index, "SP63P1")
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_SP63_P2, SummaryReserveByKey(batch, index, "SP63P2")
End Sub

' Возвращает строку данных Results, соответствующую строке LC в rngLoadCombinations.
' Если исходный номер строки не задан или вне блока, используется компактная
' раскладка по индексу результата с ограничением границами выходного массива.
Private Function OutputDataOffset(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal dataRows As Long) As Long
    OutputDataOffset = batch.SourceDataOffset(index)
    If OutputDataOffset < 1 Then OutputDataOffset = index
    If OutputDataOffset > dataRows Then OutputDataOffset = index
    If OutputDataOffset > dataRows Then OutputDataOffset = dataRows
End Function

' Возвращает запас для одного столбца компактной сводки. Для устойчивости
' ключ задает одновременно СП, главную плоскость и расчетную ветвь.
Private Function SummaryReserveByKey(ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal checkKey As String) As Double
    Select Case checkKey
        Case "StrengthStrain"
            SummaryReserveByKey = batch.ResultAt(index).StrengthResult.DirectState.StrainReserve
        Case "StrengthCapacity"
            SummaryReserveByKey = batch.ResultAt(index).StrengthResult.Capacity.ReserveFactor
        Case "Crack"
            SummaryReserveByKey = batch.ResultAt(index).CrackResult.Width.ReserveFactor
        Case "LongCrack"
            SummaryReserveByKey = batch.ResultAt(index).CrackResult.Longitudinal.ReserveFactor
        Case "SP35P1Eta"
            SummaryReserveByKey = StabilityBranchReserve(batch, index, 1, "SP35-eta")
        Case "SP35P1Table"
            SummaryReserveByKey = StabilityBranchReserve(batch, index, 1, "SP35-table")
        Case "SP35P2Eta"
            SummaryReserveByKey = StabilityBranchReserve(batch, index, 2, "SP35-eta")
        Case "SP35P2Table"
            SummaryReserveByKey = StabilityBranchReserve(batch, index, 2, "SP35-table")
        Case "SP63P1"
            SummaryReserveByKey = StabilityBranchReserve(batch, index, 1, "SP63")
        Case "SP63P2"
            SummaryReserveByKey = StabilityBranchReserve(batch, index, 2, "SP63")
    End Select
End Function

' Записывает статус одной плоскости/ветви устойчивости или N/A, если эта
' ветвь для текущего LC не применялась.
Private Sub PutStabilityBranchStatus(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal columnIndex As Long, ByVal policy As CResultStatusPolicy, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal planeIndex As Long, ByVal branchKey As String)
    If Not StabilityBranchApplies(batch, index, planeIndex, branchKey) Then
        data(rowIndex, columnIndex) = policy.NotApplicable
    ElseIf Not policy.IsFinished(MetaStatus(policy, batch.ResultAt(index).StabilityMeta)) Then
        data(rowIndex, columnIndex) = MetaStatus(policy, batch.ResultAt(index).StabilityMeta)
    ElseIf StabilityPlanePassed(batch, index, planeIndex) Then
        data(rowIndex, columnIndex) = policy.OK
    Else
        data(rowIndex, columnIndex) = policy.Fail
    End If
End Sub

' Возвращает расчетный запас одной плоскости/ветви устойчивости.
Private Function StabilityBranchReserve(ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal planeIndex As Long, ByVal branchKey As String) As Double
    If Not StabilityBranchApplies(batch, index, planeIndex, branchKey) Then Exit Function
    If planeIndex = 1 Then
        If batch.ResultAt(index).StabilityResult.Reserve1 > 0# Then StabilityBranchReserve = batch.ResultAt(index).StabilityResult.Reserve1
    Else
        If batch.ResultAt(index).StabilityResult.Reserve2 > 0# Then StabilityBranchReserve = batch.ResultAt(index).StabilityResult.Reserve2
    End If
End Function

' Проверяет, относится ли расчет устойчивости LC к конкретному столбцу
' compact summary: СП 35/СП 63, номер плоскости и ветвь table/eta.
Private Function StabilityBranchApplies(ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal planeIndex As Long, ByVal branchKey As String) As Boolean
    If Not batch.StabilityRequested(index) Then Exit Function

    Dim applicable As Boolean
    Dim actualBranch As String
    If planeIndex = 1 Then
        applicable = batch.ResultAt(index).StabilityResult.PlaneApplicable1
        actualBranch = batch.ResultAt(index).StabilityResult.PlaneBranch1
    Else
        applicable = batch.ResultAt(index).StabilityResult.PlaneApplicable2
        actualBranch = batch.ResultAt(index).StabilityResult.PlaneBranch2
    End If
    If Not applicable Then Exit Function

    If StrComp(branchKey, "SP63", vbTextCompare) = 0 Then
        StabilityBranchApplies = (StrComp(batch.ResultAt(index).StabilityResult.Code, "SP63", vbTextCompare) = 0)
    Else
        StabilityBranchApplies = (StrComp(batch.ResultAt(index).StabilityResult.Code, "SP35", vbTextCompare) = 0 And _
            StrComp(actualBranch, branchKey, vbTextCompare) = 0)
    End If
End Function

' Возвращает PASS/FAIL признак для одной плоскости устойчивости.
Private Function StabilityPlanePassed(ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal planeIndex As Long) As Boolean
    If planeIndex = 1 Then
        StabilityPlanePassed = batch.ResultAt(index).StabilityResult.PlanePassed1
    Else
        StabilityPlanePassed = batch.ResultAt(index).StabilityResult.PlanePassed2
    End If
End Function

' Записывает статус проверки или N/A, если расчетная ветка не была включена профилем.
Private Sub PutStatusCell(ByRef data() As Variant, ByVal rowIndex As Long, ByVal columnIndex As Long, _
        ByVal policy As CResultStatusPolicy, ByVal requested As Boolean, ByVal meta As CResultMeta)
    If requested Then
        data(rowIndex, columnIndex) = MetaStatus(policy, meta)
    Else
        data(rowIndex, columnIndex) = policy.NotApplicable
    End If
End Sub

' Возвращает display только через общую policy. Отсутствие обязательной meta
' является CalcErr, а не поводом угадывать статус по другому результату.
Private Function MetaStatus(ByVal policy As CResultStatusPolicy, ByVal meta As CResultMeta) As String
    MetaStatus = policy.ExternalStatus(meta)
End Function

' Возвращает пользовательское пояснение из typed-meta без анализа StopReason в writer-е.
Private Function MetaComment(ByVal meta As CResultMeta) As String
    If meta Is Nothing Then Exit Function
    MetaComment = meta.ResultComment
End Function

' Записывает только положительные расчетные коэффициенты запаса.
Private Sub PutReserveCell(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal columnIndex As Long, ByVal reserveValue As Double)
    If reserveValue > 0# Then data(rowIndex, columnIndex) = reserveValue
End Sub

' Применяет визуальное оформление compact summary: объединения, серые шапки,
' статусную палитру, единые шрифты и рамки.
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
    ' Ширина столбцов задана при сборке книги и сохраняет правки пользователя.

    stage = "HeaderStyle"
    target.Offset(0, 0).Resize(META_ROWS, 5).Interior.Color = HEADER_COLOR
    target.Offset(0, 1).Resize(META_ROWS, 4).Font.Italic = True
    target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(WORST_HEADER_ROW - GROUP_HEADER_ROW + 1, TARGET_COLUMNS).Interior.Color = HEADER_COLOR
    target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(COLUMN_HEADER_LAST_ROW - GROUP_HEADER_ROW + 1, TARGET_COLUMNS).Font.Bold = True
    target.Rows.Item(COMMENT_HEADER_ROW).Font.Bold = False
    target.Rows.Item(WORST_HEADER_ROW).Font.Bold = True

    stage = "MergeHeaders"
    MergeSummaryBlock target, 1, 1, 2, 5
    MergeSummaryBlock target, 2, 2, 2, 5
    MergeSummaryBlock target, 3, 3, 2, 5
    MergeSummaryBlock target, 4, 4, 2, 5
    MergeSummaryBlock target, GROUP_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_ID, SUM_COL_ID
    MergeSummaryBlock target, GROUP_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_COMMENT, SUM_COL_COMMENT
    MergeSummaryBlock target, GROUP_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_RESULT_COMMENT, SUM_COL_RESULT_COMMENT
    MergeSummaryBlock target, GROUP_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_OVERALL_STATUS, SUM_COL_OVERALL_STATUS
    MergeSummaryHeader target, GROUP_HEADER_ROW, SUM_COL_STATUS_STRAIN, SUM_COL_STATUS_SP63_P2
    MergeSummaryHeader target, GROUP_HEADER_ROW, SUM_COL_RESERVE_STRAIN, SUM_COL_RESERVE_SP63_P2
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, SUM_COL_STATUS_STRAIN, SUM_COL_STATUS_CAPACITY
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, SUM_COL_STATUS_CRACK, SUM_COL_STATUS_LONG_CRACK
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, SUM_COL_STATUS_SP35_P1_ETA, SUM_COL_STATUS_SP35_P2_TABLE
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, SUM_COL_STATUS_SP63_P1, SUM_COL_STATUS_SP63_P2
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_STRAIN, SUM_COL_RESERVE_CAPACITY
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_CRACK, SUM_COL_RESERVE_LONG_CRACK
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_SP35_P1_ETA, SUM_COL_RESERVE_SP35_P2_TABLE
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_SP63_P1, SUM_COL_RESERVE_SP63_P2
    MergeColumnLabels target, COLUMN_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_STATUS_STRAIN, SUM_COL_STATUS_CAPACITY
    MergeColumnLabels target, COLUMN_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_STATUS_CRACK, SUM_COL_STATUS_LONG_CRACK
    MergeColumnLabels target, COLUMN_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_RESERVE_STRAIN, SUM_COL_RESERVE_CAPACITY
    MergeColumnLabels target, COLUMN_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_RESERVE_CRACK, SUM_COL_RESERVE_LONG_CRACK
    MergeSummaryHeader target, COLUMN_HEADER_ROW, SUM_COL_STATUS_SP35_P1_ETA, SUM_COL_STATUS_SP35_P1_TABLE
    MergeSummaryHeader target, COLUMN_HEADER_ROW, SUM_COL_STATUS_SP35_P2_ETA, SUM_COL_STATUS_SP35_P2_TABLE
    MergeSummaryHeader target, COLUMN_HEADER_ROW, SUM_COL_RESERVE_SP35_P1_ETA, SUM_COL_RESERVE_SP35_P1_TABLE
    MergeSummaryHeader target, COLUMN_HEADER_ROW, SUM_COL_RESERVE_SP35_P2_ETA, SUM_COL_RESERVE_SP35_P2_TABLE
    MergeColumnLabels target, COLUMN_HEADER_ROW, COLUMN_HEADER_ROW + 1, SUM_COL_STATUS_SP63_P1, SUM_COL_STATUS_SP63_P2
    MergeColumnLabels target, COLUMN_HEADER_ROW, COLUMN_HEADER_ROW + 1, SUM_COL_RESERVE_SP63_P1, SUM_COL_RESERVE_SP63_P2

    stage = "Alignment"
    target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(WORST_HEADER_ROW - GROUP_HEADER_ROW + 1, TARGET_COLUMNS).HorizontalAlignment = -4108
    target.Rows.Item(COMMENT_HEADER_ROW).HorizontalAlignment = -4131
    target.Offset(0, 0).Resize(META_ROWS, 1).HorizontalAlignment = -4108
    target.Offset(0, 1).Resize(META_ROWS, 4).HorizontalAlignment = -4131
    target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(COLUMN_HEADER_LAST_ROW - GROUP_HEADER_ROW + 1, TARGET_COLUMNS).WrapText = True
    target.Rows.Item(COMMENT_HEADER_ROW).WrapText = False
    target.Rows.Item(WORST_HEADER_ROW).WrapText = False

    Dim dataRows As Long
    dataRows = target.Rows.Count - FIRST_DATA_ROW + 1
    If dataRows < 1 Then dataRows = 1
    target.Offset(FIRST_DATA_ROW - 1, 0).Resize(dataRows, TARGET_COLUMNS).HorizontalAlignment = -4108
    target.Offset(FIRST_DATA_ROW - 1, SUM_COL_COMMENT - 1).Resize(dataRows, 1).HorizontalAlignment = -4131
    target.Offset(FIRST_DATA_ROW - 1, SUM_COL_RESULT_COMMENT - 1).Resize(dataRows, 1).HorizontalAlignment = -4131
    target.Offset(GROUP_HEADER_ROW - 1, SUM_COL_STATUS_STRAIN - 2).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, 1).Interior.ColorIndex = -4142
    target.Offset(GROUP_HEADER_ROW - 1, SUM_COL_RESERVE_STRAIN - 2).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, 1).Interior.ColorIndex = -4142
    stage = "FormatStatusCells"
    FormatStatusCells target, batch
    stage = "WriteStatusLegend"
    WriteStatusLegend target

    stage = "Borders"
    ApplySummaryBorders target.Offset(0, 0).Resize(META_ROWS, 5)
    ApplySummaryBorders target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, SUM_COL_OVERALL_STATUS)
    ApplySummaryBorders target.Offset(GROUP_HEADER_ROW - 1, SUM_COL_STATUS_STRAIN - 1).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, _
        SUM_COL_STATUS_SP63_P2 - SUM_COL_STATUS_STRAIN + 1)
    ApplySummaryBorders target.Offset(GROUP_HEADER_ROW - 1, SUM_COL_RESERVE_STRAIN - 1).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, _
        SUM_COL_RESERVE_SP63_P2 - SUM_COL_RESERVE_STRAIN + 1)
    stage = "NumberFormat"
    target.Offset(WORST_HEADER_ROW - 1, SUM_COL_RESERVE_STRAIN - 1).Resize(target.Rows.Count - WORST_HEADER_ROW + 1, _
        SUM_COL_RESERVE_SP63_P2 - SUM_COL_RESERVE_STRAIN + 1).NumberFormat = "0.000"
    Exit Sub

Failed:
    Err.Raise Err.Number, "CBatchResultWriter.FormatSummary", stage & ": " & Err.Description
End Sub

' Пишет справа от rngBatchSummary небольшой словарь пользовательских статусов.
' Положение считается от якоря summary, поэтому при переносе rngBatchSummary
' поясняющий блок переезжает вместе с основной сводкой.
Private Sub WriteStatusLegend(ByVal target As Object)
    Dim legend As Object
    Set legend = target.Offset(GROUP_HEADER_ROW - 1, TARGET_COLUMNS + STATUS_LEGEND_GAP_COLUMNS). _
        Resize(STATUS_LEGEND_ROWS, STATUS_LEGEND_COLUMNS)

    legend.UnMerge
    legend.Clear
    legend.Font.Name = "Arial"
    legend.Font.Size = 9
    legend.VerticalAlignment = -4108
    legend.WrapText = False

    Dim data() As Variant
    ReDim data(1 To STATUS_LEGEND_ROWS, 1 To STATUS_LEGEND_COLUMNS)
    data(1, 1) = "Расшифровка статусов"
    data(2, 1) = "Статус"
    data(2, 2) = "Описание"
    data(3, 1) = "OK"
    data(3, 2) = "Расчетная ветка выполнена, проверка пройдена."
    data(4, 1) = "FAIL"
    data(4, 2) = "Равновесие найдено, но физический предел или проверка не проходят."
    data(5, 1) = "BaseFail"
    data(5, 2) = "Несущая способность не проходит уже для исходной части выбранного пути."
    data(6, 1) = "NumFail"
    data(6, 2) = "Равновесие или расчетная точка не сошлись численно."
    data(7, 1) = "InputErr"
    data(7, 2) = "Ошибка исходных данных или настройки."
    data(8, 1) = "CalcErr"
    data(8, 2) = "Внутренняя ошибка программы или несформированный результат."
    data(9, 1) = "N/A"
    data(9, 2) = "Проверка не применялась для данного сочетания."
    legend.Value2 = data

    MergeSummaryHeader legend, 1, 1, STATUS_LEGEND_COLUMNS

    legend.Rows.Item(1).Font.Bold = True
    legend.Rows.Item(2).Font.Bold = True
    legend.Offset(2, 0).Resize(STATUS_LEGEND_ROWS - 2, STATUS_LEGEND_COLUMNS).Font.Italic = True
    legend.Rows.Item(1).Interior.Color = HEADER_COLOR
    legend.Rows.Item(2).Interior.Color = HEADER_COLOR
    Dim rowIndex As Long
    For rowIndex = 3 To STATUS_LEGEND_ROWS
        ApplyStatusFill legend.Cells(rowIndex, 1), CStr(data(rowIndex, 1))
    Next rowIndex
    legend.Columns.Item(1).HorizontalAlignment = -4108
    legend.Columns.Item(2).HorizontalAlignment = -4131
    ApplySummaryBorders legend
End Sub

' Объединяет горизонтальный заголовок одной строки.
Private Sub MergeSummaryHeader(ByVal target As Object, ByVal rowIndex As Long, ByVal firstColumn As Long, ByVal lastColumn As Long)
    Dim area As Object
    Set area = target.Resize(1, 1).Offset(rowIndex - 1, firstColumn - 1). _
        Resize(1, lastColumn - firstColumn + 1)
    area.Merge
End Sub

' Объединяет прямоугольный блок шапки summary.
Private Sub MergeSummaryBlock(ByVal target As Object, ByVal firstRow As Long, ByVal lastRow As Long, _
        ByVal firstColumn As Long, ByVal lastColumn As Long)
    Dim area As Object
    Set area = target.Resize(1, 1).Offset(firstRow - 1, firstColumn - 1). _
        Resize(lastRow - firstRow + 1, lastColumn - firstColumn + 1)
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

' Наносит стандартные черные рамки для компактных блоков Results.
Private Sub ApplySummaryBorders(ByVal area As Object)
    area.Borders.LineStyle = 1
    area.Borders.Color = RGB(0, 0, 0)
    area.Borders.Weight = 2
End Sub

' Красит статусные ячейки compact summary по единой палитре OK/FAIL/BaseFail/NumFail/InputErr/CalcErr/N/A.
Private Sub FormatStatusCells(ByVal target As Object, ByVal batch As CBatchSectionCalculator)
    If batch Is Nothing Then Exit Sub
    If batch.Count <= 0 Then Exit Sub

    Dim i As Long
    For i = 1 To batch.Count
        Dim rowNumber As Long
        rowNumber = FIRST_DATA_ROW + OutputDataOffset(batch, i, target.Rows.Count - FIRST_DATA_ROW + 1) - 2
        ApplyStatusCell target, rowNumber, SUM_COL_OVERALL_STATUS
        ApplyStatusCell target, rowNumber, SUM_COL_STATUS_STRAIN
        ApplyStatusCell target, rowNumber, SUM_COL_STATUS_CAPACITY
        ApplyStatusCell target, rowNumber, SUM_COL_STATUS_CRACK
        ApplyStatusCell target, rowNumber, SUM_COL_STATUS_LONG_CRACK
        ApplyStatusCell target, rowNumber, SUM_COL_STATUS_SP35_P1_ETA
        ApplyStatusCell target, rowNumber, SUM_COL_STATUS_SP35_P1_TABLE
        ApplyStatusCell target, rowNumber, SUM_COL_STATUS_SP35_P2_ETA
        ApplyStatusCell target, rowNumber, SUM_COL_STATUS_SP35_P2_TABLE
        ApplyStatusCell target, rowNumber, SUM_COL_STATUS_SP63_P1
        ApplyStatusCell target, rowNumber, SUM_COL_STATUS_SP63_P2
    Next i
End Sub

' Красит одну статусную ячейку строки LC, не затрагивая соседние коэффициенты запаса.
Private Sub ApplyStatusCell(ByVal target As Object, ByVal rowNumber As Long, ByVal columnIndex As Long)
    Dim cell As Object
    Set cell = target.Offset(rowNumber, columnIndex - 1).Resize(1, 1)
    If Len(Trim$(CStr(cell.Value2))) = 0 Then Exit Sub
    ApplyStatusFill cell, CStr(cell.Value2)
End Sub

' Возвращает общий объект политики статусов для writer-а.
Private Function StatusPolicy() As CResultStatusPolicy
    Static policy As CResultStatusPolicy
    If policy Is Nothing Then Set policy = New CResultStatusPolicy
    Set StatusPolicy = policy
End Function

' Формирует коэффициент в заголовке с точкой как десятичным разделителем.
Private Function FormatCoefficientInvariant(ByVal value As Double) As String
    FormatCoefficientInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' Возвращает подпись запаса для eta-ветви СП35 с фактическим лимитом из Config.
Private Function SP35EtaReserveLabel(ByVal batch As CBatchSectionCalculator) As String
    Dim limitValue As Double
    limitValue = 0.7
    If Not batch Is Nothing Then
        If batch.StabilitySP35NOverNcrLimit > 0# Then limitValue = batch.StabilitySP35NOverNcrLimit
    End If
    SP35EtaReserveLabel = FormatCoefficientInvariant(limitValue) & "*Ncr/N"
End Function

' Утилитарный максимум для размеров диапазонов Results.
Private Function MaxLong(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxLong = a Else MaxLong = b
End Function
