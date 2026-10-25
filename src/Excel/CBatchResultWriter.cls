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
' Сводка пакетного расчёта на листе Results
' ==========================================================================
' Записывает rngBatchSummary блоками из общего массива по готовым данным
' CBatchSectionCalculator. В сводке показываются определяющее сочетание,
' статусы проверок и минимальные коэффициенты запаса.
' Подробные таблицы прочности, трещин и устойчивости заполняются отдельными классами.

Private mCapturedSnapshot As CCalculationSnapshot ' Приёмник общей и подробных сводок до сохранения снимка.
Private mMinimalFormatting As Boolean ' Сохраняет готовое оформление; обновляет данные, шапки при изменении и цвета статусов.

Private Const TARGET_COLUMNS As Long = 33 ' Сводка, один пустой столбец и шесть нагрузок в пользовательской точке.
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
Private Const SUM_COL_RESULT_COMMENT As Long = 3 ' Готовое пояснение итогового результата сочетания нагрузок.
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
Private Const SUM_COL_RESERVE_STRAIN As Long = 17 ' Запас по деформациям при проверке прочности.
Private Const SUM_COL_RESERVE_CAPACITY As Long = 18 ' Запас по предельным усилиям.
Private Const SUM_COL_RESERVE_CRACK As Long = 19 ' Запас по нормальным трещинам.
Private Const SUM_COL_RESERVE_LONG_CRACK As Long = 20 ' Запас по продольным трещинам.
Private Const SUM_COL_RESERVE_SP35_P1_ETA As Long = 21 ' Запас СП35, плоскость 1, ec > r: limit*Ncr/N.
Private Const SUM_COL_RESERVE_SP35_P1_TABLE As Long = 22 ' Запас СП35, плоскость 1, ec <= r: Nult/N.
Private Const SUM_COL_RESERVE_SP35_P2_ETA As Long = 23 ' Запас СП35, плоскость 2, ec > r: limit*Ncr/N.
Private Const SUM_COL_RESERVE_SP35_P2_TABLE As Long = 24 ' Запас СП35, плоскость 2, ec <= r: Nult/N.
Private Const SUM_COL_RESERVE_SP63_P1 As Long = 25 ' Запас СП63, плоскость 1: Ncr/N.
Private Const SUM_COL_RESERVE_SP63_P2 As Long = 26 ' Запас СП63, плоскость 2: Ncr/N.
Private Const SUM_COL_ORIGINAL_N As Long = 28 ' Исходная N до фильтров и продольного изгиба.
Private Const SUM_COL_ORIGINAL_MX As Long = 29 ' Исходный Mx в пользовательской точке.
Private Const SUM_COL_ORIGINAL_MY As Long = 30 ' Исходный My в пользовательской точке.
Private Const SUM_COL_DESIGN_N As Long = 31 ' Рабочая N, не умножаемая на коэффициент eta.
Private Const SUM_COL_DESIGN_MX As Long = 32 ' Рабочий Mx после продольного изгиба, если он применим.
Private Const SUM_COL_DESIGN_MY As Long = 33 ' Рабочий My после продольного изгиба, если он применим.

Private Const STATUS_LEGEND_GAP_COLUMNS As Long = 1 ' Один пустой столбец между общей сводкой и расшифровкой статусов.
Private Const STATUS_LEGEND_COLUMNS As Long = 2 ' Ширина правого словаря статусов: код и текстовая расшифровка.
Private Const STATUS_LEGEND_ROWS As Long = 9 ' Заголовок, шапка и семь пользовательских статусов.

Private Const DEFAULT_SUMMARY_ROWS As Long = 42 ' Штатная высота шаблона: 12 строк до данных и 30 строк LC.

Private Const HEADER_COLOR As Long = &HD9D9D9 ' Светло-серый фон заголовков.

' Возвращает число строк общей сводки для заданного количества сочетаний.
Public Function RequiredSummaryRows(ByVal combinationCount As Long) As Long
    If combinationCount < 0 Then combinationCount = 0
    RequiredSummaryRows = FIRST_DATA_ROW - 1 + combinationCount
    If RequiredSummaryRows < COLUMN_HEADER_ROW Then RequiredSummaryRows = COLUMN_HEADER_ROW
    If RequiredSummaryRows < DEFAULT_SUMMARY_ROWS Then RequiredSummaryRows = DEFAULT_SUMMARY_ROWS
End Function

' Возвращает число строк LC по высоте rngLoadCombinations. Это нужно для
' пустых строк общей сводки: пользователь видит все доступные строки сочетаний,
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

' Возвращает число строк общей сводки с учётом доступных строк rngLoadCombinations.
Public Function RequiredSummaryRowsForWorkbook(ByVal workbook As Object, ByVal combinationCount As Long) As Long
    RequiredSummaryRowsForWorkbook = RequiredSummaryRows(RequiredDataRowsForWorkbook(workbook, combinationCount))
End Function

' Возвращает ширину собственно rngBatchSummary.
Public Function RequiredSummaryColumns() As Long
    RequiredSummaryColumns = TARGET_COLUMNS
End Function

' Возвращает число столбцов Results, занимаемых общей сводкой.
Public Function RequiredSummaryOutputColumns() As Long
    RequiredSummaryOutputColumns = TARGET_COLUMNS + STATUS_LEGEND_GAP_COLUMNS + STATUS_LEGEND_COLUMNS
End Function

' Возвращает полную высоту, которую занимает компактная сводка.
Public Function RequiredSummaryOutputRows(ByVal combinationCount As Long) As Long
    RequiredSummaryOutputRows = RequiredSummaryRows(combinationCount)
End Function

' Возвращает число строк сводки для проверки её размещения на листе Results.
Public Function RequiredSummaryOutputRowsForWorkbook(ByVal workbook As Object, ByVal combinationCount As Long) As Long
    RequiredSummaryOutputRowsForWorkbook = RequiredSummaryRowsForWorkbook(workbook, combinationCount)
End Function

' Записывает готовые результаты в rngBatchSummary блочными присваиваниями
' и оформляет сводку в выбранном полном или минимальном режиме.
' Подробные таблицы заполняют отдельные классы. Модель section передаётся
' таблице прочности для геометрических величин текущего и предельного НДС;
' сама сводка не выполняет инженерные расчёты.
Public Sub WriteSummary(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, _
        Optional ByVal units As CUnitSystem = Nothing, Optional ByVal section As CSectionModel = Nothing, Optional ByVal fullFormatting As Boolean = True)
    mMinimalFormatting = Not fullFormatting
    On Error GoTo Failed

    Dim stage As String
    stage = "ValidateArguments"
    If workbook Is Nothing Then Err.Raise vbObjectError + 3970, "CBatchResultWriter", "Книга Excel не передана."
    If batch Is Nothing Then Err.Raise vbObjectError + 3971, "CBatchResultWriter", "Пакетный расчетчик не передан."

    Dim summaryRows As Long
    summaryRows = RequiredSummaryRowsForWorkbook(workbook, batch.Count)

    stage = "PrepareSummaryTarget"
    If mCapturedSnapshot Is Nothing Then PrepareSummaryTarget SummaryClearRange(workbook), summaryRows

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
    If batch.WorstCombinationIndex > 0 Then PutPointLoads data, WORST_HEADER_ROW, batch, batch.WorstCombinationIndex, units

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
Public Sub ClearSummary(ByVal workbook As Object, Optional ByVal fullFormatting As Boolean = True)
    mMinimalFormatting = Not fullFormatting
    If workbook Is Nothing Then Err.Raise vbObjectError + 3972, "CBatchResultWriter", "Книга Excel не передана."
    ClearSummaryBody SummaryClearRange(workbook)
    ClearStrengthSummary workbook
    ClearCrackSummary workbook
    ClearStabilitySummary workbook
End Sub

' После записи Results и построения схемы дописывает полное время запуска.
' Сама сводка готовится раньше подробных таблиц, поскольку схема
' Plot.LoadCase = Worst читает из неё определяющее сочетание.
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

' Возвращает область очистки общей сводки до заголовка следующей таблицы.
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
    If strengthTopRow > anchor.Row Then
        Dim gap As Long: gap = strengthTopRow - anchor.Row
        If mMinimalFormatting Then gap = gap - 1
        rowsToClear = MaxLong(rowsToClear, gap)
    End If

UseDefaultHeight:
    If anchor.Row + rowsToClear - 1 > anchor.Worksheet.Rows.Count Then rowsToClear = anchor.Worksheet.Rows.Count - anchor.Row + 1
    Set SummaryClearRange = anchor.Resize(rowsToClear, RequiredSummaryOutputColumns())
End Function

' При полном оформлении снимает объединения и очищает старую сводку.
' В минимальном режиме очищает содержимое, сохраняя готовые форматы ячеек.
Private Sub PrepareSummaryTarget(ByVal target As Object, ByVal requiredRows As Long)
    If target.Rows.Count < requiredRows Then Set target = target.Resize(requiredRows, target.Columns.Count)
    ClearSummaryBody target
End Sub

' Очищает данные прежней сводки. В минимальном режиме сохраняет шапку
' и основные форматы, но убирает прежнюю подсветку статусов.
Private Sub ClearSummaryBody(ByVal target As Object)
    If mMinimalFormatting Then
        target.Cells(1, 1).Resize(META_ROWS, 1).ClearContents
        target.Cells(1, 7).Resize(3, 1).ClearContents
        target.Offset(WORST_HEADER_ROW - 1, 0).Resize(target.Rows.Count - WORST_HEADER_ROW + 1, TARGET_COLUMNS).ClearContents
        If target.Rows.Count >= FIRST_DATA_ROW Then target.Offset(FIRST_DATA_ROW - 1, 0).Resize(target.Rows.Count - FIRST_DATA_ROW + 1, TARGET_COLUMNS).Interior.ColorIndex = -4142
        Exit Sub
    End If
    target.UnMerge
    target.Clear
End Sub

' Записывает сводку блоками из массива VBA. В минимальном режиме сохраняет
' совпадающую шапку и обновляет три участка данных; изменённую шапку создаёт заново.
Private Sub WriteSummaryBody(ByVal target As Object, ByRef data() As Variant)
    If Not mCapturedSnapshot Is Nothing Then
        NormalizeSnapshotHeaders data
        mCapturedSnapshot.AddTable "Summary", data
        Exit Sub
    End If
    If mMinimalFormatting Then
        If ResultHeaderMatches(target, data, GROUP_HEADER_ROW, COMMENT_HEADER_ROW, merges:=SummaryHeaderBlocks()) And _
                CStr(target.Cells(target.Rows.Count, SUM_COL_RESERVE_STRAIN).NumberFormat) = "0.000" And _
                CStr(target.Cells(target.Rows.Count, SUM_COL_ORIGINAL_N).NumberFormat) = "0.000" Then
            WriteResultDataSlice target, data, 1, 1, META_ROWS, 1
            WriteResultDataSlice target, data, 1, 7, 3, 1
            WriteResultDataSlice target, data, WORST_HEADER_ROW, 1, UBound(data, 1) - WORST_HEADER_ROW + 1, TARGET_COLUMNS
            Exit Sub
        End If
        mMinimalFormatting = False
        target.UnMerge
    End If
    target.Resize(UBound(data, 1), TARGET_COLUMNS).Value2 = data
End Sub

' Сохраняет расположение надписей после стандартного Merge, не обращаясь к листу.
' Excel переносит единственное непустое значение блока в его верхнюю левую ячейку.
Private Sub NormalizeSnapshotHeaders(ByRef data() As Variant)
    Dim block As Variant, row As Long, column As Long, value As Variant, found As Boolean
    For Each block In SummaryHeaderBlocks()
        found = False: value = Empty
        For row = CLng(block(0)) To CLng(block(1))
            For column = CLng(block(2)) To CLng(block(3))
                If Not found And Len(CStr(data(row, column))) > 0 Then value = data(row, column): found = True
                data(row, column) = Empty
            Next column
        Next row
        data(CLng(block(0)), CLng(block(2))) = value
    Next block
End Sub

' Обновляет подробную таблицу устойчивости рядом с сохранёнными расчётными данными.
Private Sub WriteStabilitySummary(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    Dim writer As CStabilitySummaryWriter
    Set writer = New CStabilitySummaryWriter
    Set writer.CaptureSnapshot = mCapturedSnapshot
    writer.WriteSummary workbook, batch, units, Not mMinimalFormatting
End Sub

' Обновляет подробную таблицу НДС по прочности и несущей способности.
Private Sub WriteStrengthSummary(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, _
        ByVal units As CUnitSystem, ByVal section As CSectionModel)
    Dim writer As CStrengthSummaryWriter
    Set writer = New CStrengthSummaryWriter
    Set writer.CaptureSnapshot = mCapturedSnapshot
    writer.WriteSummary workbook, batch, units, section, Not mMinimalFormatting
End Sub

' Обновляет отдельную детальную таблицу нормальных и продольных трещин.
Private Sub WriteCrackSummary(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    Dim writer As CCrackSummaryWriter
    Set writer = New CCrackSummaryWriter
    Set writer.CaptureSnapshot = mCapturedSnapshot
    writer.WriteSummary workbook, batch, units, Not mMinimalFormatting
End Sub

' Очищает отдельный блок прочности при общей очистке Results.
Private Sub ClearStrengthSummary(ByVal workbook As Object)
    Dim writer As CStrengthSummaryWriter
    Set writer = New CStrengthSummaryWriter
    writer.ClearSummary workbook, Not mMinimalFormatting
End Sub

' Очищает отдельный блок устойчивости при общей очистке Results.
Private Sub ClearStabilitySummary(ByVal workbook As Object)
    Dim writer As CStabilitySummaryWriter
    Set writer = New CStabilitySummaryWriter
    writer.ClearSummary workbook, Not mMinimalFormatting
End Sub

' Очищает отдельный блок трещин при общей очистке Results.
Private Sub ClearCrackSummary(ByVal workbook As Object)
    Dim writer As CCrackSummaryWriter
    Set writer = New CCrackSummaryWriter
    writer.ClearSummary workbook, Not mMinimalFormatting
End Sub

' Заполняет четыре верхние строки общей сводки: определяющее сочетание, комментарий,
' число вызовов CSectionSolver, стартов UltimateStrain, итерации и время.
Private Sub PutMetaRows(ByRef data() As Variant, ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    data(1, 1) = batch.WorstCombinationID
    data(1, 2) = "Определяющее сочетание"
    data(2, 1) = batch.WorstCombinationName
    data(2, 2) = "Комментарий определяющего сочетания"
    data(3, 1) = batch.SolverCallCount
    data(3, 2) = "Количество решений НДС"
    data(4, 1) = batch.ElapsedSeconds
    data(4, 2) = "Время расчета, с"
    data(1, 7) = batch.UltimateStrainStartCount
    data(1, 8) = "Количество стартов UltimateStrain"
    data(2, 7) = batch.SolverIterationCount
    data(2, 8) = "Количество итераций НДС"
    data(3, 7) = batch.UltimateStrainIterationCount
    data(3, 8) = "Количество итераций UltimateStrain"
End Sub

' Заполняет верхнюю шапку компактной сводки: отдельные части для статусов
' проверок и коэффициентов запаса.
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

' Заполняет строки обозначений и пояснений к каждому столбцу общей сводки.
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
    Dim forceUnit As String, momentUnit As String
    forceUnit = "N": momentUnit = "N*mm"
    If Not units Is Nothing Then forceUnit = units.OutputForceUnit: momentUnit = units.OutputMomentUnit
    PutPointLoadHeaders data, forceUnit, momentUnit
End Sub

' Заполняет шестистрочную шапку нагрузок; единицы принадлежат сохранённому расчёту.
Private Sub PutPointLoadHeaders(ByRef data() As Variant, ByVal forceUnit As String, ByVal momentUnit As String)
    data(GROUP_HEADER_ROW, SUM_COL_ORIGINAL_N) = "Нагрузки в точке приложения"
    data(COLUMN_HEADER_ROW, SUM_COL_ORIGINAL_N) = "Исходные заданные нагрузки"
    data(COLUMN_HEADER_ROW, SUM_COL_DESIGN_N) = "Нагрузки с учётом продольного изгиба" & vbLf & "(если применимо)"
    data(COLUMN_HEADER_LAST_ROW, SUM_COL_ORIGINAL_N) = "N, " & forceUnit
    data(COLUMN_HEADER_LAST_ROW, SUM_COL_ORIGINAL_MX) = "Mx, " & momentUnit
    data(COLUMN_HEADER_LAST_ROW, SUM_COL_ORIGINAL_MY) = "My, " & momentUnit
    data(COLUMN_HEADER_LAST_ROW, SUM_COL_DESIGN_N) = "N, " & forceUnit
    data(COLUMN_HEADER_LAST_ROW, SUM_COL_DESIGN_MX) = "Mx, " & momentUnit
    data(COLUMN_HEADER_LAST_ROW, SUM_COL_DESIGN_MY) = "My, " & momentUnit
    data(COMMENT_HEADER_ROW, SUM_COL_ORIGINAL_N) = "Исходная продольная сила"
    data(COMMENT_HEADER_ROW, SUM_COL_ORIGINAL_MX) = "Исходный изгибающий момент Mx"
    data(COMMENT_HEADER_ROW, SUM_COL_ORIGINAL_MY) = "Исходный изгибающий момент My"
    data(COMMENT_HEADER_ROW, SUM_COL_DESIGN_N) = "Продольная сила"
    data(COMMENT_HEADER_ROW, SUM_COL_DESIGN_MX) = WorkingMomentComment("Mx")
    data(COMMENT_HEADER_ROW, SUM_COL_DESIGN_MY) = WorkingMomentComment("My")
End Sub

' Возвращает отдельный комментарий момента; продольная сила на eta не умножается.
Private Function WorkingMomentComment(ByVal component As String) As String
    WorkingMomentComment = "Изгибающий момент " & component & " с учётом " & ChrW$(&H3B7) & " от продольного изгиба (если применимо)"
End Function

' Записывает исходные и рабочие усилия из готового пакета, без обращения к Config.
' Ошибка устойчивости не подменяется якобы успешно уточнёнными моментами.
Private Sub PutPointLoads(ByRef data() As Variant, ByVal row As Long, ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal units As CUnitSystem)
    If Not batch.OriginalLoadsAvailable(index) Then Exit Sub
    data(row, SUM_COL_ORIGINAL_N) = PointLoadToOutput(batch.OriginalN(index), 0, units)
    data(row, SUM_COL_ORIGINAL_MX) = PointLoadToOutput(batch.OriginalMx(index), 1, units)
    data(row, SUM_COL_ORIGINAL_MY) = PointLoadToOutput(batch.OriginalMy(index), 2, units)
    Dim policy As CResultStatusPolicy: Set policy = StatusPolicy()
    Dim status As String: status = policy.ExternalStatus(batch.ResultAt(index).StabilityMeta)
    If status <> policy.OK And status <> policy.NotApplicable Then Exit Sub
    data(row, SUM_COL_DESIGN_N) = PointLoadToOutput(batch.N(index), 0, units)
    data(row, SUM_COL_DESIGN_MX) = PointLoadToOutput(batch.UserMx(index), 1, units)
    data(row, SUM_COL_DESIGN_MY) = PointLoadToOutput(batch.UserMy(index), 2, units)
End Sub

' Применяет единственные рабочие преобразования единиц и знаков CUnitSystem.
' component: 0 — продольная сила, 1 — Mx, 2 — My в одной точке приложения.
Private Function PointLoadToOutput(ByVal value As Double, ByVal component As Long, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then PointLoadToOutput = value: Exit Function
    Select Case component
        Case 0: PointLoadToOutput = units.InternalForceToOutput(value)
        Case 1: PointLoadToOutput = units.InternalMomentMxToOutput(value)
        Case 2: PointLoadToOutput = units.InternalMomentMyToOutput(value)
    End Select
End Function

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

' Записывает строку сочетания в общую сводку. CResultStatusPolicy задаёт
' обозначения статусов, а численные коэффициенты берутся из подробных
' результатов того же сочетания, показанных в таблицах ниже.
Private Sub PutCombination(ByRef data() As Variant, ByVal rowIndex As Long, ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal units As CUnitSystem, ByVal section As CSectionModel)
    Dim policy As CResultStatusPolicy
    Set policy = StatusPolicy()

    If rowIndex > UBound(data, 1) Then Exit Sub

    data(rowIndex, SUM_COL_ID) = batch.CombinationID(index)
    data(rowIndex, SUM_COL_COMMENT) = batch.CombinationName(index)
    data(rowIndex, SUM_COL_RESULT_COMMENT) = CalculationResultMessage(batch, index, batch.ResultAt(index).OverallMeta)
    data(rowIndex, SUM_COL_OVERALL_STATUS) = MetaStatus(policy, batch.ResultAt(index).OverallMeta)
    PutPointLoads data, rowIndex, batch, index, units

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
    ElseIf Not policy.IsMetaFinished(batch.ResultAt(index).StabilityMeta) Then
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

' Проверяет, в какой столбец общей сводки записать результат устойчивости:
' учитывает нормативный документ, главную плоскость и способ проверки.
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

' Возвращает обозначение статуса по общим правилам. Если обязательного
' объекта статуса нет, сообщает CalcErr без догадок по соседним результатам.
Private Function MetaStatus(ByVal policy As CResultStatusPolicy, ByVal meta As CResultMeta) As String
    MetaStatus = policy.ExternalStatus(meta)
End Function

' Возвращает готовое пояснение результата. Причина остановки численного

' Записывает только положительный расчётный коэффициент запаса.
' При нулевом или отрицательном значении оставляет подготовленную ячейку без изменения.
Private Sub PutReserveCell(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal columnIndex As Long, ByVal reserveValue As Double)
    If reserveValue > 0# Then data(rowIndex, columnIndex) = reserveValue
End Sub

' При полном оформлении задаёт шапку, рамки и числовые форматы сводки.
' В минимальном режиме обновляет только подсветку статусов; ширины столбцов не меняет.
Private Sub FormatSummary(ByVal workbook As Object, ByVal target As Object, ByVal batch As CBatchSectionCalculator, _
        Optional ByRef preparedValues As Variant)
    If Not mCapturedSnapshot Is Nothing Then Exit Sub
    If mMinimalFormatting Then
        FormatStatusCells target, batch, preparedValues
        Exit Sub
    End If
    On Error GoTo Failed

    Dim stage As String
    stage = "BaseFont"
    target.HorizontalAlignment = -4131
    target.VerticalAlignment = -4108
    target.WrapText = False
    target.Font.Name = "Arial"
    target.Font.Size = 9

    stage = "HeaderStyle"
    target.Offset(0, 0).Resize(META_ROWS, 5).Interior.Color = HEADER_COLOR
    target.Offset(0, 1).Resize(META_ROWS, 4).Font.Italic = True
    target.Offset(0, 6).Resize(3, 7).Interior.Color = HEADER_COLOR
    target.Offset(0, 7).Resize(3, 6).Font.Italic = True
    target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(WORST_HEADER_ROW - GROUP_HEADER_ROW + 1, TARGET_COLUMNS).Interior.Color = HEADER_COLOR
    target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(COLUMN_HEADER_LAST_ROW - GROUP_HEADER_ROW + 1, TARGET_COLUMNS).Font.Bold = True
    target.Rows.Item(COMMENT_HEADER_ROW).Font.Bold = False
    target.Rows.Item(WORST_HEADER_ROW).Font.Bold = True

    stage = "MergeHeaders"
    Dim block As Variant
    For Each block In SummaryHeaderBlocks()
        MergeSummaryBlock target, CLng(block(0)), CLng(block(1)), CLng(block(2)), CLng(block(3))
    Next block

    stage = "Alignment"
    target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(WORST_HEADER_ROW - GROUP_HEADER_ROW + 1, TARGET_COLUMNS).HorizontalAlignment = -4108
    target.Rows.Item(COMMENT_HEADER_ROW).HorizontalAlignment = -4131
    target.Offset(0, 0).Resize(META_ROWS, 1).HorizontalAlignment = -4108
    target.Offset(0, 1).Resize(META_ROWS, 4).HorizontalAlignment = -4131
    target.Offset(0, 6).Resize(3, 1).HorizontalAlignment = -4108
    target.Offset(0, 6).Resize(3, 1).NumberFormat = "0"
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
    target.Offset(GROUP_HEADER_ROW - 1, SUM_COL_ORIGINAL_N - 2).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, 1).Interior.ColorIndex = -4142
    stage = "FormatStatusCells"
    FormatStatusCells target, batch, preparedValues
    stage = "WriteStatusLegend"
    WriteStatusLegend target

    stage = "Borders"
    ApplySummaryBorders target.Offset(0, 0).Resize(META_ROWS, 5)
    ApplySummaryBorders target.Offset(0, 6).Resize(3, 7)
    ApplySummaryBorders target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, SUM_COL_OVERALL_STATUS)
    ApplySummaryBorders target.Offset(GROUP_HEADER_ROW - 1, SUM_COL_STATUS_STRAIN - 1).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, _
        SUM_COL_STATUS_SP63_P2 - SUM_COL_STATUS_STRAIN + 1)
    ApplySummaryBorders target.Offset(GROUP_HEADER_ROW - 1, SUM_COL_RESERVE_STRAIN - 1).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, _
        SUM_COL_RESERVE_SP63_P2 - SUM_COL_RESERVE_STRAIN + 1)
    ApplySummaryBorders target.Offset(GROUP_HEADER_ROW - 1, SUM_COL_ORIGINAL_N - 1).Resize(target.Rows.Count - GROUP_HEADER_ROW + 1, _
        SUM_COL_DESIGN_MY - SUM_COL_ORIGINAL_N + 1)
    stage = "NumberFormat"
    target.Offset(WORST_HEADER_ROW - 1, SUM_COL_RESERVE_STRAIN - 1).Resize(target.Rows.Count - WORST_HEADER_ROW + 1, _
        SUM_COL_RESERVE_SP63_P2 - SUM_COL_RESERVE_STRAIN + 1).NumberFormat = "0.000"
    target.Offset(WORST_HEADER_ROW - 1, SUM_COL_ORIGINAL_N - 1).Resize(target.Rows.Count - WORST_HEADER_ROW + 1, _
        SUM_COL_DESIGN_MY - SUM_COL_ORIGINAL_N + 1).NumberFormat = "0.000"
    Exit Sub

Failed:
    Err.Raise Err.Number, "CBatchResultWriter.FormatSummary", stage & ": " & Err.Description
End Sub

' Пишет пояснения пользовательских статусов справа от rngBatchSummary.
' Положение отсчитывается от начала сводки и меняется вместе с её переносом.
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
    data(5, 2) = "Исходная постоянная часть выбранного пути уже не проходит проверку."
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

' Описывает объединения один раз для проверки готовой шапки и её построения.
Private Function SummaryHeaderBlocks() As Collection
    Dim blocks As Collection: Set blocks = New Collection
    Dim column As Long
    blocks.Add Array(1, 1, 2, 5)
    blocks.Add Array(2, 2, 2, 5)
    blocks.Add Array(3, 3, 2, 5)
    blocks.Add Array(4, 4, 2, 5)
    blocks.Add Array(1, 1, 8, 13)
    blocks.Add Array(2, 2, 8, 13)
    blocks.Add Array(3, 3, 8, 13)
    blocks.Add Array(GROUP_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_ID, SUM_COL_ID)
    blocks.Add Array(GROUP_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_COMMENT, SUM_COL_COMMENT)
    blocks.Add Array(GROUP_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_RESULT_COMMENT, SUM_COL_RESULT_COMMENT)
    blocks.Add Array(GROUP_HEADER_ROW, COLUMN_HEADER_LAST_ROW, SUM_COL_OVERALL_STATUS, SUM_COL_OVERALL_STATUS)
    blocks.Add Array(GROUP_HEADER_ROW, GROUP_HEADER_ROW, SUM_COL_STATUS_STRAIN, SUM_COL_STATUS_SP63_P2)
    blocks.Add Array(GROUP_HEADER_ROW, GROUP_HEADER_ROW, SUM_COL_RESERVE_STRAIN, SUM_COL_RESERVE_SP63_P2)
    blocks.Add Array(GROUP_HEADER_ROW, SUBGROUP_HEADER_ROW, SUM_COL_ORIGINAL_N, SUM_COL_DESIGN_MY)
    blocks.Add Array(COLUMN_HEADER_ROW, COLUMN_HEADER_ROW + 1, SUM_COL_ORIGINAL_N, SUM_COL_ORIGINAL_MY)
    blocks.Add Array(COLUMN_HEADER_ROW, COLUMN_HEADER_ROW + 1, SUM_COL_DESIGN_N, SUM_COL_DESIGN_MY)
    blocks.Add Array(SUBGROUP_HEADER_ROW, SUBGROUP_HEADER_ROW, SUM_COL_STATUS_STRAIN, SUM_COL_STATUS_CAPACITY)
    blocks.Add Array(SUBGROUP_HEADER_ROW, SUBGROUP_HEADER_ROW, SUM_COL_STATUS_CRACK, SUM_COL_STATUS_LONG_CRACK)
    blocks.Add Array(SUBGROUP_HEADER_ROW, SUBGROUP_HEADER_ROW, SUM_COL_STATUS_SP35_P1_ETA, SUM_COL_STATUS_SP35_P2_TABLE)
    blocks.Add Array(SUBGROUP_HEADER_ROW, SUBGROUP_HEADER_ROW, SUM_COL_STATUS_SP63_P1, SUM_COL_STATUS_SP63_P2)
    blocks.Add Array(SUBGROUP_HEADER_ROW, SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_STRAIN, SUM_COL_RESERVE_CAPACITY)
    blocks.Add Array(SUBGROUP_HEADER_ROW, SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_CRACK, SUM_COL_RESERVE_LONG_CRACK)
    blocks.Add Array(SUBGROUP_HEADER_ROW, SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_SP35_P1_ETA, SUM_COL_RESERVE_SP35_P2_TABLE)
    blocks.Add Array(SUBGROUP_HEADER_ROW, SUBGROUP_HEADER_ROW, SUM_COL_RESERVE_SP63_P1, SUM_COL_RESERVE_SP63_P2)
    For column = SUM_COL_STATUS_STRAIN To SUM_COL_STATUS_CAPACITY
        blocks.Add Array(COLUMN_HEADER_ROW, COLUMN_HEADER_LAST_ROW, column, column)
    Next column
    For column = SUM_COL_STATUS_CRACK To SUM_COL_STATUS_LONG_CRACK
        blocks.Add Array(COLUMN_HEADER_ROW, COLUMN_HEADER_LAST_ROW, column, column)
    Next column
    For column = SUM_COL_RESERVE_STRAIN To SUM_COL_RESERVE_CAPACITY
        blocks.Add Array(COLUMN_HEADER_ROW, COLUMN_HEADER_LAST_ROW, column, column)
    Next column
    For column = SUM_COL_RESERVE_CRACK To SUM_COL_RESERVE_LONG_CRACK
        blocks.Add Array(COLUMN_HEADER_ROW, COLUMN_HEADER_LAST_ROW, column, column)
    Next column
    blocks.Add Array(COLUMN_HEADER_ROW, COLUMN_HEADER_ROW, SUM_COL_STATUS_SP35_P1_ETA, SUM_COL_STATUS_SP35_P1_TABLE)
    blocks.Add Array(COLUMN_HEADER_ROW, COLUMN_HEADER_ROW, SUM_COL_STATUS_SP35_P2_ETA, SUM_COL_STATUS_SP35_P2_TABLE)
    blocks.Add Array(COLUMN_HEADER_ROW, COLUMN_HEADER_ROW, SUM_COL_RESERVE_SP35_P1_ETA, SUM_COL_RESERVE_SP35_P1_TABLE)
    blocks.Add Array(COLUMN_HEADER_ROW, COLUMN_HEADER_ROW, SUM_COL_RESERVE_SP35_P2_ETA, SUM_COL_RESERVE_SP35_P2_TABLE)
    For column = SUM_COL_STATUS_SP63_P1 To SUM_COL_STATUS_SP63_P2
        blocks.Add Array(COLUMN_HEADER_ROW, COLUMN_HEADER_ROW + 1, column, column)
    Next column
    For column = SUM_COL_RESERVE_SP63_P1 To SUM_COL_RESERVE_SP63_P2
        blocks.Add Array(COLUMN_HEADER_ROW, COLUMN_HEADER_ROW + 1, column, column)
    Next column
    Set SummaryHeaderBlocks = blocks
End Function

' Объединяет горизонтальный заголовок одной строки.
Private Sub MergeSummaryHeader(ByVal target As Object, ByVal rowIndex As Long, ByVal firstColumn As Long, ByVal lastColumn As Long)
    Dim area As Object
    Set area = target.Resize(1, 1).Offset(rowIndex - 1, firstColumn - 1). _
        Resize(1, lastColumn - firstColumn + 1)
    area.Merge
End Sub

' Объединяет прямоугольный блок заголовка общей сводки.
Private Sub MergeSummaryBlock(ByVal target As Object, ByVal firstRow As Long, ByVal lastRow As Long, _
        ByVal firstColumn As Long, ByVal lastColumn As Long)
    Dim area As Object
    Set area = target.Resize(1, 1).Offset(firstRow - 1, firstColumn - 1). _
        Resize(lastRow - firstRow + 1, lastColumn - firstColumn + 1)
    area.Merge
End Sub



' Наносит стандартные черные рамки для компактных блоков Results.
Private Sub ApplySummaryBorders(ByVal area As Object)
    area.Borders.LineStyle = 1
    area.Borders.Color = RGB(0, 0, 0)
    area.Borders.Weight = 2
End Sub

' Задаёт цвет ячеек статусов общей сводки по единой палитре modStatusFormatting.
Private Sub FormatStatusCells(ByVal target As Object, ByVal batch As CBatchSectionCalculator, Optional ByRef preparedValues As Variant)
    Dim rows() As Long, i As Long
    If batch Is Nothing Then
        Dim storedRows As Variant: storedRows = ResultSnapshotStatusRows(target, FIRST_DATA_ROW - 1, preparedValues, FIRST_DATA_ROW - 1)
        If IsEmpty(storedRows) Then Exit Sub
        rows = storedRows
    Else
    If batch.Count <= 0 Then Exit Sub
    ReDim rows(1 To batch.Count)
    For i = 1 To batch.Count
        rows(i) = FIRST_DATA_ROW + OutputDataOffset(batch, i, target.Rows.Count - FIRST_DATA_ROW + 1) - 1
    Next i
    End If
    ApplyTableStatusFills target, rows, Array( _
        SUM_COL_OVERALL_STATUS, SUM_COL_STATUS_STRAIN, SUM_COL_STATUS_CAPACITY, _
        SUM_COL_STATUS_CRACK, SUM_COL_STATUS_LONG_CRACK, SUM_COL_STATUS_SP35_P1_ETA, _
        SUM_COL_STATUS_SP35_P1_TABLE, SUM_COL_STATUS_SP35_P2_ETA, SUM_COL_STATUS_SP35_P2_TABLE, _
        SUM_COL_STATUS_SP63_P1, SUM_COL_STATUS_SP63_P2), True, 0, preparedValues
End Sub


' Возвращает общие правила вывода статусов для записи результатов.
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

' Передаёт отдельный приёмник готовых массивов; при Nothing сохраняется обычный вывод.
Friend Property Set CaptureSnapshot(ByVal value As CCalculationSnapshot)
    Set mCapturedSnapshot = value
End Property

' Восстанавливает общую и три подробные сводки из одних сохранённых массивов.
Public Sub RestoreSummary(ByVal workbook As Object, ByVal snapshot As CCalculationSnapshot, Optional ByVal fullFormatting As Boolean = True)
    Set mCapturedSnapshot = Nothing
    mMinimalFormatting = Not fullFormatting
    Dim data() As Variant: data = SummaryForDisplay(snapshot)
    PrepareSummaryTarget SummaryClearRange(workbook), UBound(data, 1)
    Dim target As Object: Set target = SummaryTargetRange(workbook, UBound(data, 1))
    WriteSummaryBody target, data
    FormatSummary workbook, target, Nothing, data
    Dim strength As CStrengthSummaryWriter: Set strength = New CStrengthSummaryWriter
    Dim cracks As CCrackSummaryWriter: Set cracks = New CCrackSummaryWriter
    Dim stability As CStabilitySummaryWriter: Set stability = New CStabilitySummaryWriter
    strength.RestoreSummary workbook, snapshot, fullFormatting
    cracks.RestoreSummary workbook, snapshot, fullFormatting
    stability.RestoreSummary workbook, snapshot, fullFormatting
End Sub

' Читает прежние файловые снимки без блока нагрузок, не придумывая исходный ввод.
' Это совместимость внешнего формата: текущий Config и содержимое файла не меняются.
Private Function SummaryForDisplay(ByVal snapshot As CCalculationSnapshot) As Variant
    Dim saved As Variant: saved = snapshot.Table("Summary")
    If UBound(saved, 2) = TARGET_COLUMNS Then
        saved(COMMENT_HEADER_ROW, SUM_COL_DESIGN_MX) = WorkingMomentComment("Mx")
        saved(COMMENT_HEADER_ROW, SUM_COL_DESIGN_MY) = WorkingMomentComment("My")
        SummaryForDisplay = saved: Exit Function
    End If
    Dim data() As Variant, row As Long, column As Long
    ReDim data(1 To UBound(saved, 1), 1 To TARGET_COLUMNS)
    For row = 1 To UBound(saved, 1)
        For column = 1 To UBound(saved, 2): data(row, column) = saved(row, column): Next column
    Next row
    PutPointLoadHeaders data, SavedSummaryUnit(snapshot, "Output.ForceUnit"), SavedSummaryUnit(snapshot, "Output.MomentUnit")
    For column = SUM_COL_ORIGINAL_N To SUM_COL_ORIGINAL_MY
        data(COMMENT_HEADER_ROW, column) = "Не сохранено в этом снимке"
    Next column
    data(COMMENT_HEADER_ROW, SUM_COL_DESIGN_N) = "Не сохранено в этом снимке"
    data(COMMENT_HEADER_ROW, SUM_COL_DESIGN_MX) = "Блок нагрузок отсутствует в прежнем снимке"
    data(COMMENT_HEADER_ROW, SUM_COL_DESIGN_MY) = "Блок нагрузок отсутствует в прежнем снимке"
    SummaryForDisplay = data
End Function

' Получает единицу прежнего файлового снимка, независимо от текущих настроек.
Private Function SavedSummaryUnit(ByVal snapshot As CCalculationSnapshot, ByVal key As String) As String
    Dim data As Variant, row As Long: data = snapshot.Table("rngNDMSectionProperties")
    For row = 2 To UBound(data, 1)
        If CStr(data(row, 3)) = key Then SavedSummaryUnit = CStr(data(row, 4)): Exit Function
    Next row
    Err.Raise 5, "CBatchResultWriter", "В снимке отсутствует единица " & key
End Function
