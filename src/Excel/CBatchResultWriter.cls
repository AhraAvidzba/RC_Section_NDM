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

Private Const TARGET_COLUMNS As Long = 25 ' Ширина верхней сводки с разложением устойчивости по СП, плоскостям и ветвям.
Private Const LEGACY_SUMMARY_CLEAR_COLUMNS As Long = 34 ' Чистит старую широкую сводку и правый словарь статусов при первом запуске новой версии.
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
Private Const SUM_COL_OVERALL_STATUS As Long = 3 ' Итоговый статус LC по всем включенным проверкам.
Private Const SUM_COL_STATUS_STRAIN As Long = 5 ' Статус прямого НДС по модели прочности.
Private Const SUM_COL_STATUS_CAPACITY As Long = 6 ' Статус проверки несущей способности.
Private Const SUM_COL_STATUS_CRACK As Long = 7 ' Статус нормальных трещин.
Private Const SUM_COL_STATUS_LONG_CRACK As Long = 8 ' Статус продольных трещин.
Private Const SUM_COL_STATUS_SP35_P1_ETA As Long = 9 ' СП35, плоскость 1, ветвь ec > r.
Private Const SUM_COL_STATUS_SP35_P1_TABLE As Long = 10 ' СП35, плоскость 1, ветвь ec <= r.
Private Const SUM_COL_STATUS_SP35_P2_ETA As Long = 11 ' СП35, плоскость 2, ветвь ec > r.
Private Const SUM_COL_STATUS_SP35_P2_TABLE As Long = 12 ' СП35, плоскость 2, ветвь ec <= r.
Private Const SUM_COL_STATUS_SP63_P1 As Long = 13 ' СП63, плоскость 1.
Private Const SUM_COL_STATUS_SP63_P2 As Long = 14 ' СП63, плоскость 2.
Private Const SUM_COL_RESERVE_STRAIN As Long = 16 ' Запас по деформациям strength-модели.
Private Const SUM_COL_RESERVE_CAPACITY As Long = 17 ' Запас по предельным усилиям.
Private Const SUM_COL_RESERVE_CRACK As Long = 18 ' Запас по нормальным трещинам.
Private Const SUM_COL_RESERVE_LONG_CRACK As Long = 19 ' Запас по продольным трещинам.
Private Const SUM_COL_RESERVE_SP35_P1_ETA As Long = 20 ' Запас СП35, плоскость 1, ec > r: limit*Ncr/N.
Private Const SUM_COL_RESERVE_SP35_P1_TABLE As Long = 21 ' Запас СП35, плоскость 1, ec <= r: Nult/N.
Private Const SUM_COL_RESERVE_SP35_P2_ETA As Long = 22 ' Запас СП35, плоскость 2, ec > r: limit*Ncr/N.
Private Const SUM_COL_RESERVE_SP35_P2_TABLE As Long = 23 ' Запас СП35, плоскость 2, ec <= r: Nult/N.
Private Const SUM_COL_RESERVE_SP63_P1 As Long = 24 ' Запас СП63, плоскость 1: Ncr/N.
Private Const SUM_COL_RESERVE_SP63_P2 As Long = 25 ' Запас СП63, плоскость 2: Ncr/N.

Private Const DEFAULT_SUMMARY_ROWS As Long = 42 ' Штатная высота шаблона: 12 строк до данных и 30 строк LC.
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
    PutColumnHeaders data, units, batch
    stage = "PutWorstRow"
    PutWorstRow data, batch

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

' Возвращает область rngBatchSummary нужного размера и сразу проверяет, хватает ли листа Results.
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
    data(COLUMN_HEADER_ROW + 1, sp63FirstColumn) = "Ncr/N"
    data(COLUMN_HEADER_ROW + 1, sp63FirstColumn + 1) = "Ncr/N"
End Sub

' Заполняет строку "worst LC" отдельно по каждой проверке. В левом блоке
' выводится номер сочетания, где запас минимален, а в правом блоке - сам
' этот минимальный запас. Так сводка сразу показывает, какое LC управляет
' каждой конкретной проверкой.
Private Sub PutWorstRow(ByRef data() As Variant, ByVal batch As CBatchSectionCalculator)
    Dim policy As CBatchStatusPolicy
    Set policy = StatusPolicy()

    data(WORST_HEADER_ROW, SUM_COL_ID) = "worst LC"
    PutWorstCheck data, batch, policy, SUM_COL_STATUS_STRAIN, SUM_COL_RESERVE_STRAIN, "StrengthStrain"
    PutWorstCheck data, batch, policy, SUM_COL_STATUS_CAPACITY, SUM_COL_RESERVE_CAPACITY, "StrengthCapacity"
    PutWorstCheck data, batch, policy, SUM_COL_STATUS_CRACK, SUM_COL_RESERVE_CRACK, "Crack"
    PutWorstCheck data, batch, policy, SUM_COL_STATUS_LONG_CRACK, SUM_COL_RESERVE_LONG_CRACK, "LongCrack"
    PutWorstCheck data, batch, policy, SUM_COL_STATUS_SP35_P1_ETA, SUM_COL_RESERVE_SP35_P1_ETA, "SP35P1Eta"
    PutWorstCheck data, batch, policy, SUM_COL_STATUS_SP35_P1_TABLE, SUM_COL_RESERVE_SP35_P1_TABLE, "SP35P1Table"
    PutWorstCheck data, batch, policy, SUM_COL_STATUS_SP35_P2_ETA, SUM_COL_RESERVE_SP35_P2_ETA, "SP35P2Eta"
    PutWorstCheck data, batch, policy, SUM_COL_STATUS_SP35_P2_TABLE, SUM_COL_RESERVE_SP35_P2_TABLE, "SP35P2Table"
    PutWorstCheck data, batch, policy, SUM_COL_STATUS_SP63_P1, SUM_COL_RESERVE_SP63_P1, "SP63P1"
    PutWorstCheck data, batch, policy, SUM_COL_STATUS_SP63_P2, SUM_COL_RESERVE_SP63_P2, "SP63P2"
End Sub

' Выбирает худшее LC для одного столбца проверки по минимальному положительному
' запасу. Неприменимые ветви и пустые запасы не участвуют в выборе.
Private Sub PutWorstCheck(ByRef data() As Variant, ByVal batch As CBatchSectionCalculator, _
        ByVal policy As CBatchStatusPolicy, ByVal statusColumn As Long, ByVal reserveColumn As Long, _
        ByVal checkKey As String)
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
    Else
        data(WORST_HEADER_ROW, statusColumn) = policy.NotApplicable
        data(WORST_HEADER_ROW, reserveColumn) = policy.NotApplicable
    End If
End Sub

' Записывает одну строку сочетания в compact summary. Все статусы идут через
' CBatchStatusPolicy, а коэффициенты берутся из подробных расчетных результатов
' batch, чтобы summary совпадала с детальными таблицами ниже.
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
    PutStabilityBranchStatus data, rowIndex, SUM_COL_STATUS_SP35_P1_ETA, policy, batch, index, 1, "SP35-eta"
    PutStabilityBranchStatus data, rowIndex, SUM_COL_STATUS_SP35_P1_TABLE, policy, batch, index, 1, "SP35-table"
    PutStabilityBranchStatus data, rowIndex, SUM_COL_STATUS_SP35_P2_ETA, policy, batch, index, 2, "SP35-eta"
    PutStabilityBranchStatus data, rowIndex, SUM_COL_STATUS_SP35_P2_TABLE, policy, batch, index, 2, "SP35-table"
    PutStabilityBranchStatus data, rowIndex, SUM_COL_STATUS_SP63_P1, policy, batch, index, 1, "SP63"
    PutStabilityBranchStatus data, rowIndex, SUM_COL_STATUS_SP63_P2, policy, batch, index, 2, "SP63"

    PutReserveCell data, rowIndex, SUM_COL_RESERVE_STRAIN, batch.StrengthStrainReserve(index)
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_CAPACITY, batch.StrengthCapacityReserve(index)
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_CRACK, batch.CrackReserveFactor(index)
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_LONG_CRACK, batch.LongitudinalCrackReserveFactor(index)
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_SP35_P1_ETA, SummaryReserveByKey(batch, index, "SP35P1Eta")
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_SP35_P1_TABLE, SummaryReserveByKey(batch, index, "SP35P1Table")
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_SP35_P2_ETA, SummaryReserveByKey(batch, index, "SP35P2Eta")
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_SP35_P2_TABLE, SummaryReserveByKey(batch, index, "SP35P2Table")
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_SP63_P1, SummaryReserveByKey(batch, index, "SP63P1")
    PutReserveCell data, rowIndex, SUM_COL_RESERVE_SP63_P2, SummaryReserveByKey(batch, index, "SP63P2")
End Sub

' Возвращает запас для одного столбца компактной сводки. Для устойчивости
' ключ задает одновременно СП, главную плоскость и расчетную ветвь.
Private Function SummaryReserveByKey(ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal checkKey As String) As Double
    Select Case checkKey
        Case "StrengthStrain"
            SummaryReserveByKey = batch.StrengthStrainReserve(index)
        Case "StrengthCapacity"
            SummaryReserveByKey = batch.StrengthCapacityReserve(index)
        Case "Crack"
            SummaryReserveByKey = batch.CrackReserveFactor(index)
        Case "LongCrack"
            SummaryReserveByKey = batch.LongitudinalCrackReserveFactor(index)
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
        ByVal columnIndex As Long, ByVal policy As CBatchStatusPolicy, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal planeIndex As Long, ByVal branchKey As String)
    If Not StabilityBranchApplies(batch, index, planeIndex, branchKey) Then
        data(rowIndex, columnIndex) = policy.NotApplicable
    ElseIf Not policy.IsFinished(batch.StabilityStatus(index)) Then
        data(rowIndex, columnIndex) = policy.DisplayStatus(batch.StabilityStatus(index))
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
        If batch.StabilityReserve1(index) > 0# Then StabilityBranchReserve = batch.StabilityReserve1(index)
    Else
        If batch.StabilityReserve2(index) > 0# Then StabilityBranchReserve = batch.StabilityReserve2(index)
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
        applicable = batch.StabilityPlaneApplicable1(index)
        actualBranch = batch.StabilityPlaneBranch1(index)
    Else
        applicable = batch.StabilityPlaneApplicable2(index)
        actualBranch = batch.StabilityPlaneBranch2(index)
    End If
    If Not applicable Then Exit Function

    If StrComp(branchKey, "SP63", vbTextCompare) = 0 Then
        StabilityBranchApplies = (StrComp(batch.StabilityCode(index), "SP63", vbTextCompare) = 0)
    Else
        StabilityBranchApplies = (StrComp(batch.StabilityCode(index), "SP35", vbTextCompare) = 0 And _
            StrComp(actualBranch, branchKey, vbTextCompare) = 0)
    End If
End Function

' Возвращает PASS/FAIL признак для одной плоскости устойчивости.
Private Function StabilityPlanePassed(ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal planeIndex As Long) As Boolean
    If planeIndex = 1 Then
        StabilityPlanePassed = batch.StabilityPlanePassed1(index)
    Else
        StabilityPlanePassed = batch.StabilityPlanePassed2(index)
    End If
End Function

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

' Применяет визуальное оформление compact summary: объединения, серые шапки,
' подсветку failed-статусов, единые шрифты и рамки.
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
    MergeColumnLabels target, COLUMN_HEADER_ROW + 1, COLUMN_HEADER_LAST_ROW, SUM_COL_STATUS_SP63_P1, SUM_COL_STATUS_SP63_P2
    MergeColumnLabels target, COLUMN_HEADER_ROW + 1, COLUMN_HEADER_LAST_ROW, SUM_COL_RESERVE_SP63_P1, SUM_COL_RESERVE_SP63_P2

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
    target.Offset(WORST_HEADER_ROW - 1, SUM_COL_RESERVE_STRAIN - 1).Resize(target.Rows.Count - WORST_HEADER_ROW + 1, _
        SUM_COL_RESERVE_SP63_P2 - SUM_COL_RESERVE_STRAIN + 1).Interior.Color = SAFETY_COLOR
    target.Offset(GROUP_HEADER_ROW - 1, 3).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, 1).Interior.ColorIndex = -4142
    target.Offset(GROUP_HEADER_ROW - 1, 14).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, 1).Interior.ColorIndex = -4142
    stage = "FormatOverallStatusRows"
    FormatOverallStatusRows target, batch

    stage = "Borders"
    ApplySummaryBorders target.Offset(0, 0).Resize(META_ROWS, 5)
    ApplySummaryBorders target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, 3)
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

' Объединяет горизонтальный заголовок одной строки.
Private Sub MergeSummaryHeader(ByVal target As Object, ByVal rowIndex As Long, ByVal firstColumn As Long, ByVal lastColumn As Long)
    Dim area As Object
    Set area = target.Offset(rowIndex - 1, firstColumn - 1).Resize(1, lastColumn - firstColumn + 1)
    area.Merge
End Sub

' Объединяет прямоугольный блок шапки summary.
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

' Наносит стандартные черные рамки для компактных блоков Results.
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
        HighlightStabilityFailedPair target, rowNumber, SUM_COL_STATUS_SP35_P1_ETA, SUM_COL_RESERVE_SP35_P1_ETA, batch, i, 1, "SP35-eta"
        HighlightStabilityFailedPair target, rowNumber, SUM_COL_STATUS_SP35_P1_TABLE, SUM_COL_RESERVE_SP35_P1_TABLE, batch, i, 1, "SP35-table"
        HighlightStabilityFailedPair target, rowNumber, SUM_COL_STATUS_SP35_P2_ETA, SUM_COL_RESERVE_SP35_P2_ETA, batch, i, 2, "SP35-eta"
        HighlightStabilityFailedPair target, rowNumber, SUM_COL_STATUS_SP35_P2_TABLE, SUM_COL_RESERVE_SP35_P2_TABLE, batch, i, 2, "SP35-table"
        HighlightStabilityFailedPair target, rowNumber, SUM_COL_STATUS_SP63_P1, SUM_COL_RESERVE_SP63_P1, batch, i, 1, "SP63"
        HighlightStabilityFailedPair target, rowNumber, SUM_COL_STATUS_SP63_P2, SUM_COL_RESERVE_SP63_P2, batch, i, 2, "SP63"
    Next i
End Sub

' Подсвечивает только ту ветвь устойчивости, которая реально применялась к LC.
Private Sub HighlightStabilityFailedPair(ByVal target As Object, ByVal rowNumber As Long, _
        ByVal statusColumn As Long, ByVal reserveColumn As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal planeIndex As Long, ByVal branchKey As String)
    If Not StabilityBranchApplies(batch, index, planeIndex, branchKey) Then Exit Sub
    If StabilityPlanePassed(batch, index, planeIndex) Then Exit Sub
    HighlightCell target, rowNumber, statusColumn
    HighlightCell target, rowNumber, reserveColumn
End Sub

' Подсвечивает пару "статус + запас", если соответствующая проверка не прошла.
Private Sub HighlightFailedPair(ByVal target As Object, ByVal rowNumber As Long, _
        ByVal statusColumn As Long, ByVal reserveColumn As Long, _
        ByVal policy As CBatchStatusPolicy, ByVal statusText As String)
    If Not policy.IsFailedOverall(statusText) Then Exit Sub
    HighlightCell target, rowNumber, statusColumn
    HighlightCell target, rowNumber, reserveColumn
End Sub

' Подсвечивает одиночный статус, если CBatchStatusPolicy считает его failing.
Private Sub HighlightFailedCell(ByVal target As Object, ByVal rowNumber As Long, _
        ByVal columnIndex As Long, ByVal policy As CBatchStatusPolicy, ByVal statusText As String)
    If policy.IsFailedOverall(statusText) Then HighlightCell target, rowNumber, columnIndex
End Sub

' Применяет цветовую подсветку к одной ячейке summary.
Private Sub HighlightCell(ByVal target As Object, ByVal rowNumber As Long, ByVal columnIndex As Long)
    With target.Offset(rowNumber, columnIndex - 1).Resize(1, 1)
        .Interior.Color = FAILED_ROW_COLOR
        .Font.Color = FAILED_FONT_COLOR
    End With
End Sub

' Возвращает общий объект политики статусов для writer-а.
Private Function StatusPolicy() As CBatchStatusPolicy
    Static policy As CBatchStatusPolicy
    If policy Is Nothing Then Set policy = New CBatchStatusPolicy
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
