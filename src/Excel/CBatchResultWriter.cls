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
' Writer формирует верхний диапазон rngBatchSummary как компактный инженерный
' отчет по всем сочетаниям. Таблица разделена на прямое НДС, расчет
' несущей способности и раскрытие трещин. Расчетные алгоритмы здесь не
' запускаются: класс только берет готовые данные CBatchSectionCalculator и
' записывает их на лист одним массивом, а коэффициенты запаса и a_crc задает
' формулами Excel, чтобы пользователь видел связь между исходными колонками.

Private Const TARGET_ROWS As Long = 29 ' Высота фиксированного диапазона batch summary: 9 строк шапки + 20 LC.
Private Const TARGET_COLUMNS As Long = 99 ' Ширина сводки до общего минимального запаса.
Private Const META_ROWS As Long = 6 ' Верхние строки с общими сведениями о запуске.
Private Const GROUP_HEADER_ROW As Long = 7 ' Строка объединенных заголовков расчетных блоков.
Private Const SUBGROUP_HEADER_ROW As Long = 8 ' Строка внутренних подзаголовков расчетных блоков.
Private Const COLUMN_HEADER_ROW As Long = 9 ' Строка заголовков конкретных параметров.
Private Const FIRST_DATA_ROW As Long = 10 ' Первая строка сочетаний.

Private Const COL_OVERALL_STATUS As Long = 1
Private Const COL_ID As Long = 2
Private Const COL_COMMENT As Long = 3
Private Const COL_PROFILE_ID As Long = 4

Private Const COL_DIRECT_STATUS As Long = 5
Private Const COL_EPS_C_COMP As Long = 6
Private Const COL_EPS_C_TENS As Long = 7
Private Const COL_EPS_S_COMP As Long = 8
Private Const COL_EPS_S_TENS As Long = 9
Private Const COL_STATE_H As Long = 10
Private Const COL_STATE_C As Long = 11
Private Const COL_LIM_C_COMP As Long = 12
Private Const COL_LIM_C_TENS As Long = 13
Private Const COL_LIM_S_COMP As Long = 14
Private Const COL_LIM_S_TENS As Long = 15

Private Const COL_CAPACITY_STATUS As Long = 16
Private Const COL_CAPACITY_LIMIT_STATE As Long = 17
Private Const COL_CAPACITY_LOAD_PATH As Long = 18
Private Const COL_CAPACITY_SOLUTION_METHOD As Long = 19
Private Const COL_CAPACITY_LAMBDA As Long = 20
Private Const COL_CAPACITY_N_ULT As Long = 21
Private Const COL_CAPACITY_MX_ULT As Long = 22
Private Const COL_CAPACITY_MY_ULT As Long = 23
Private Const COL_CAPACITY_H As Long = 24
Private Const COL_CAPACITY_C As Long = 25
Private Const COL_CAPACITY_SF As Long = 26

Private Const COL_CRACK_STATUS As Long = 27
Private Const COL_CRACK_FORMED As Long = 28
Private Const COL_CRACK_H As Long = 29
Private Const COL_CRACK_A As Long = 30
Private Const COL_CRACK_TENSION_DEPTH As Long = 31
Private Const COL_CRACK_EFFECTIVE_DEPTH As Long = 32
Private Const COL_CRACK_AS As Long = 33
Private Const COL_CRACK_ABT As Long = 34
Private Const COL_CRACK_DS As Long = 35
Private Const COL_CRACK_REBAR_COUNT As Long = 36
Private Const COL_CRACK_LS_RAW As Long = 37
Private Const COL_CRACK_LS As Long = 38
Private Const COL_CRACK_PHI1 As Long = 39
Private Const COL_CRACK_PHI2 As Long = 40
Private Const COL_CRACK_PHI3 As Long = 41
Private Const COL_CRACK_NCRC As Long = 42
Private Const COL_CRACK_LAMBDA_CRC As Long = 43
Private Const COL_CRACK_SIGMA_S As Long = 44
Private Const COL_CRACK_SIGMA_S_CRC As Long = 45
Private Const COL_CRACK_PSI As Long = 46
Private Const COL_CRACK_ES As Long = 47
Private Const COL_CRACK_WIDTH As Long = 48
Private Const COL_CRACK_ALLOWABLE As Long = 49
Private Const COL_CRACK_SF As Long = 50

Private Const COL_LONG_CRACK_STATUS As Long = 51
Private Const COL_LONG_CRACK_SIGMA_C_MAX As Long = 52
Private Const COL_LONG_CRACK_RB_MC2 As Long = 53
Private Const COL_LONG_CRACK_SF As Long = 54

Private Const COL_STABILITY_STATUS As Long = 55
Private Const COL_STABILITY_NCR_MIN As Long = 56
Private Const COL_STABILITY_RESERVE As Long = 57
Private Const COL_STABILITY_DESIGN_N As Long = 58
Private Const COL_STABILITY_LOADPOINT_MX As Long = 59
Private Const COL_STABILITY_LOADPOINT_MY As Long = 60
Private Const COL_STABILITY_DESIGN_MX As Long = 61
Private Const COL_STABILITY_DESIGN_MY As Long = 62

Private Const COL_STABILITY_SP35_NCR1 As Long = 63
Private Const COL_STABILITY_SP35_NULT1 As Long = 64
Private Const COL_STABILITY_SP35_RESERVE1 As Long = 65
Private Const COL_STABILITY_SP35_ETA1 As Long = 66
Private Const COL_STABILITY_SP35_PHIL1 As Long = 67
Private Const COL_STABILITY_SP35_E1 As Long = 68
Private Const COL_STABILITY_SP35_EA1 As Long = 69
Private Const COL_STABILITY_SP35_L01 As Long = 70
Private Const COL_STABILITY_SP35_D1 As Long = 71
Private Const COL_STABILITY_SP35_NCR2 As Long = 72
Private Const COL_STABILITY_SP35_NULT2 As Long = 73
Private Const COL_STABILITY_SP35_RESERVE2 As Long = 74
Private Const COL_STABILITY_SP35_ETA2 As Long = 75
Private Const COL_STABILITY_SP35_PHIL2 As Long = 76
Private Const COL_STABILITY_SP35_E2 As Long = 77
Private Const COL_STABILITY_SP35_EA2 As Long = 78
Private Const COL_STABILITY_SP35_L02 As Long = 79
Private Const COL_STABILITY_SP35_D2 As Long = 80

Private Const COL_STABILITY_SP63_NCR1 As Long = 81
Private Const COL_STABILITY_SP63_NULT1 As Long = 82
Private Const COL_STABILITY_SP63_RESERVE1 As Long = 83
Private Const COL_STABILITY_SP63_ETA1 As Long = 84
Private Const COL_STABILITY_SP63_PHIL1 As Long = 85
Private Const COL_STABILITY_SP63_E1 As Long = 86
Private Const COL_STABILITY_SP63_EA1 As Long = 87
Private Const COL_STABILITY_SP63_L01 As Long = 88
Private Const COL_STABILITY_SP63_D1 As Long = 89
Private Const COL_STABILITY_SP63_NCR2 As Long = 90
Private Const COL_STABILITY_SP63_NULT2 As Long = 91
Private Const COL_STABILITY_SP63_RESERVE2 As Long = 92
Private Const COL_STABILITY_SP63_ETA2 As Long = 93
Private Const COL_STABILITY_SP63_PHIL2 As Long = 94
Private Const COL_STABILITY_SP63_E2 As Long = 95
Private Const COL_STABILITY_SP63_EA2 As Long = 96
Private Const COL_STABILITY_SP63_L02 As Long = 97
Private Const COL_STABILITY_SP63_D2 As Long = 98

Private Const COL_OVERALL_SF As Long = 99

Private Const HEADER_COLOR As Long = &HD9D9D9 ' Светло-серый фон заголовков.
Private Const SAFETY_COLOR As Long = &HE7E6E6 ' Серый фон столбцов коэффициентов запаса.
Private Const OVERALL_SAFETY_COLOR As Long = &HBFBFBF ' Более темный фон итогового минимального запаса.
Private Const FAILED_ROW_COLOR As Long = &HCEC7FF ' Светло-красный фон LC, где расчет не прошел или не сошелся.
Private Const FAILED_FONT_COLOR As Long = &H9C0006 ' Темно-красный текст для строк FAIL/NumFail/InputErr.
Private Const NOT_CALCULATED_SHORT As String = "NotCalcul." ' Короткий статус для компактной таблицы Results.

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

    stage = "SummaryTargetRange"
    Dim target As Object
    Set target = SummaryTargetRange(workbook)
    stage = "PrepareSummaryTarget"
    PrepareSummaryTarget target

    stage = "AllocateData"
    Dim data() As Variant
    ReDim data(1 To TARGET_ROWS, 1 To TARGET_COLUMNS)

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
    Exit Sub

Failed:
    Err.Raise Err.Number, "CBatchResultWriter.WriteSummary", stage & ": " & Err.Description
End Sub

' Очищает текущую сводку перед новым расчетом или ручной очисткой Results.
Public Sub ClearSummary(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3972, "CBatchResultWriter", "Книга Excel не передана."
    ClearSummaryBody SummaryTargetRange(workbook)
End Sub

' Обновляет итоговое время уже после завершения записи Results и схемы.
' Сводка формируется перед детальным Results, потому что Plot.LoadCase = Worst читает из
' нее определяющее сочетание; поэтому общий elapsed time дозаписывается в конце.
Public Sub UpdateElapsedSeconds(ByVal workbook As Object, ByVal elapsedSeconds As Double)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3974, "CBatchResultWriter", "Книга Excel не передана."

    Dim target As Object
    Set target = SummaryTargetRange(workbook)
    target.Offset(4, 4).Resize(1, 1).Value2 = elapsedSeconds
End Sub

Private Function SummaryTargetRange(ByVal workbook As Object) As Object
    Dim target As Object
    Set target = workbook.Names.Item("rngBatchSummary").RefersToRange
    If target.Rows.Count < TARGET_ROWS Or target.Columns.Count < TARGET_COLUMNS Then
        Err.Raise vbObjectError + 3973, "CBatchResultWriter", "rngBatchSummary меньше требуемого размера сводки."
    End If
    Set SummaryTargetRange = target.Resize(TARGET_ROWS, TARGET_COLUMNS)
End Function

' Снимает объединения и очищает текущую сводку до блочной записи массива.
' Это важно при повторном расчете: предыдущий запуск объединяет строки
' заголовков, и Excel не позволяет записывать двумерный массив поверх
' частично объединенного диапазона.
Private Sub PrepareSummaryTarget(ByVal target As Object)
    ClearSummaryBody target
End Sub

Private Sub ClearSummaryBody(ByVal target As Object)
    Dim body As Object
    Set body = target.Offset(1, 0).Resize(TARGET_ROWS - 1, TARGET_COLUMNS)
    body.UnMerge
    body.Clear
End Sub

Private Sub WriteSummaryBody(ByVal target As Object, ByRef data() As Variant)
    Dim body() As Variant
    ReDim body(1 To TARGET_ROWS - 1, 1 To TARGET_COLUMNS)

    Dim rowIndex As Long
    Dim columnIndex As Long
    For rowIndex = 2 To TARGET_ROWS
        For columnIndex = 1 To TARGET_COLUMNS
            body(rowIndex - 1, columnIndex) = data(rowIndex, columnIndex)
        Next columnIndex
    Next rowIndex

    target.Offset(1, 0).Resize(TARGET_ROWS - 1, TARGET_COLUMNS).Formula = body
End Sub

Private Sub PutMetaRows(ByRef data() As Variant, ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    data(2, 1) = "Определяющее сочетание по прочности"
    data(2, 5) = batch.GoverningCombinationID
    data(3, 1) = "Комментарий определяющего сочетания"
    data(3, 5) = batch.GoverningCombinationName
    data(4, 1) = "Определяющее сочетание по трещинам"
    data(4, 5) = batch.CrackGoverningCombinationID
    data(4, 8) = batch.CrackGoverningCombinationName
    data(5, 1) = "Время расчета, с"
    data(5, 5) = batch.ElapsedSeconds
    data(6, 1) = "Точка приложения нагрузки (относительно центра тяжести бетонного сечения)"
    data(6, 5) = "X=" & CStr(OutputLengthForSummary(batch.LoadReferenceOffsetX, units)) & " " & OutputLengthUnit(units)
    data(6, 8) = "Y=" & CStr(OutputLengthForSummary(batch.LoadReferenceOffsetY, units)) & " " & OutputLengthUnit(units)
End Sub

Private Sub PutGroupHeaders(ByRef data() As Variant)
    data(GROUP_HEADER_ROW, COL_DIRECT_STATUS) = "Прямое состояние НДС и физические пределы"
    data(GROUP_HEADER_ROW, COL_CAPACITY_STATUS) = "Расчет предельной несущей способности по выбранной " & ChrW$(&H3BB) & "-траектории"
    data(GROUP_HEADER_ROW, COL_CRACK_STATUS) = "Расчет ширины раскрытия нормальных трещин"
    data(GROUP_HEADER_ROW, COL_LONG_CRACK_STATUS) = "Продольные трещины"
    data(GROUP_HEADER_ROW, COL_STABILITY_STATUS) = "Продольный изгиб и устойчивость"
    data(GROUP_HEADER_ROW, COL_OVERALL_SF) = "Минимальный запас"
End Sub

Private Sub PutSubgroupHeaders(ByRef data() As Variant)
    data(SUBGROUP_HEADER_ROW, COL_OVERALL_STATUS) = "Сочетание"
    data(SUBGROUP_HEADER_ROW, COL_DIRECT_STATUS) = "Фактическое НДС от LC"
    data(SUBGROUP_HEADER_ROW, COL_LIM_C_COMP) = "Предельные деформации"

    data(SUBGROUP_HEADER_ROW, COL_CAPACITY_STATUS) = "Статус, путь и метод"
    data(SUBGROUP_HEADER_ROW, COL_CAPACITY_LAMBDA) = "Множитель"
    data(SUBGROUP_HEADER_ROW, COL_CAPACITY_N_ULT) = "Предельные компоненты"
    data(SUBGROUP_HEADER_ROW, COL_CAPACITY_H) = "Геометрия при " & ChrW$(&H3BB) & "ult"
    data(SUBGROUP_HEADER_ROW, COL_CAPACITY_SF) = "Запас"

    data(SUBGROUP_HEADER_ROW, COL_CRACK_STATUS) = "Фактическое НДС от LC"
    data(SUBGROUP_HEADER_ROW, COL_CRACK_AS) = "Арматура и бетон"
    data(SUBGROUP_HEADER_ROW, COL_CRACK_LS_RAW) = "Расстояние между трещинами"
    data(SUBGROUP_HEADER_ROW, COL_CRACK_PHI1) = "Коэффициенты phi"
    data(SUBGROUP_HEADER_ROW, COL_CRACK_NCRC) = "Расчет psi_s"
    data(SUBGROUP_HEADER_ROW, COL_CRACK_ES) = "Модуль арматуры"
    data(SUBGROUP_HEADER_ROW, COL_CRACK_WIDTH) = "Проверка a_crc"
    data(SUBGROUP_HEADER_ROW, COL_LONG_CRACK_STATUS) = "Проверка сжатого бетона"
    data(SUBGROUP_HEADER_ROW, COL_STABILITY_STATUS) = "Итог и усилия после фильтра"
    data(SUBGROUP_HEADER_ROW, COL_STABILITY_SP35_NCR1) = "СП 35 - плоскость 1"
    data(SUBGROUP_HEADER_ROW, COL_STABILITY_SP35_NCR2) = "СП 35 - плоскость 2"
    data(SUBGROUP_HEADER_ROW, COL_STABILITY_SP63_NCR1) = "СП 63 - плоскость 1"
    data(SUBGROUP_HEADER_ROW, COL_STABILITY_SP63_NCR2) = "СП 63 - плоскость 2"
    data(SUBGROUP_HEADER_ROW, COL_OVERALL_SF) = "Итог"
End Sub

Private Sub PutColumnHeaders(ByRef data() As Variant, ByVal units As CUnitSystem)
    data(COLUMN_HEADER_ROW, COL_OVERALL_STATUS) = "OverallStatus (итог)"
    data(COLUMN_HEADER_ROW, COL_ID) = "CombinationID (сочетание)"
    data(COLUMN_HEADER_ROW, COL_COMMENT) = "Comment (комментарий)"
    data(COLUMN_HEADER_ROW, COL_PROFILE_ID) = "ProfileId (профиль)"
    data(COLUMN_HEADER_ROW, COL_DIRECT_STATUS) = "DirectStateStatus (НДС)"

    data(COLUMN_HEADER_ROW, COL_EPS_C_COMP) = "eps_c,min (бетон сжатие)"
    data(COLUMN_HEADER_ROW, COL_EPS_C_TENS) = "eps_c,max (бетон растяжение)"
    data(COLUMN_HEADER_ROW, COL_EPS_S_COMP) = "eps_s,min (арматура сжатие)"
    data(COLUMN_HEADER_ROW, COL_EPS_S_TENS) = "eps_s,max (арматура растяжение)"
    data(COLUMN_HEADER_ROW, COL_STATE_H) = "h, " & OutputLengthUnit(units) & " (высота текущего НДС)"
    data(COLUMN_HEADER_ROW, COL_STATE_C) = "x, " & OutputLengthUnit(units) & " (сжатая зона текущего НДС)"
    data(COLUMN_HEADER_ROW, COL_LIM_C_COMP) = "eps_cu,comp (предел бетона сжатие)"
    data(COLUMN_HEADER_ROW, COL_LIM_C_TENS) = "eps_cu,tens (предел бетона растяжение)"
    data(COLUMN_HEADER_ROW, COL_LIM_S_COMP) = "eps_su,comp (предел арматуры сжатие)"
    data(COLUMN_HEADER_ROW, COL_LIM_S_TENS) = "eps_su,tens (предел арматуры растяжение)"
    data(COLUMN_HEADER_ROW, COL_CAPACITY_STATUS) = "CapacityStatus (прочность)"
    data(COLUMN_HEADER_ROW, COL_CAPACITY_LIMIT_STATE) = "CapacityLimitState (предельное состояние)"
    data(COLUMN_HEADER_ROW, COL_CAPACITY_LOAD_PATH) = "CapacityLoadPath (что масштабируется)"
    data(COLUMN_HEADER_ROW, COL_CAPACITY_SOLUTION_METHOD) = "CapacitySolutionMethod (фактический метод)"
    data(COLUMN_HEADER_ROW, COL_CAPACITY_LAMBDA) = "lambdaUltimate (множитель)"
    data(COLUMN_HEADER_ROW, COL_CAPACITY_N_ULT) = "Nult, " & OutputForceUnit(units) & " (предельная N)"
    data(COLUMN_HEADER_ROW, COL_CAPACITY_MX_ULT) = "Mxult, " & OutputMomentUnit(units) & " (предельный Mx)"
    data(COLUMN_HEADER_ROW, COL_CAPACITY_MY_ULT) = "Myult, " & OutputMomentUnit(units) & " (предельный My)"
    data(COLUMN_HEADER_ROW, COL_CAPACITY_H) = "h, " & OutputLengthUnit(units) & " (высота при lambdaUltimate)"
    data(COLUMN_HEADER_ROW, COL_CAPACITY_C) = "x, " & OutputLengthUnit(units) & " (сжатая зона при lambdaUltimate)"
    data(COLUMN_HEADER_ROW, COL_CAPACITY_SF) = "CapacitySafetyFactor (запас по " & ChrW$(&H3BB) & ")"

    data(COLUMN_HEADER_ROW, COL_CRACK_STATUS) = "CrackStatus (статус трещин)"
    data(COLUMN_HEADER_ROW, COL_CRACK_FORMED) = "CrackFormed (трещина есть)"
    data(COLUMN_HEADER_ROW, COL_CRACK_H) = "h, " & OutputLengthUnit(units) & " (высота сечения)"
    data(COLUMN_HEADER_ROW, COL_CRACK_A) = "a, " & OutputLengthUnit(units) & " (от края до стержня)"
    data(COLUMN_HEADER_ROW, COL_CRACK_TENSION_DEPTH) = "TensionZoneDepth, " & OutputLengthUnit(units) & " (растянутая зона)"
    data(COLUMN_HEADER_ROW, COL_CRACK_EFFECTIVE_DEPTH) = "EffectiveZoneDepth, " & OutputLengthUnit(units) & " (эффективная зона)"
    data(COLUMN_HEADER_ROW, COL_CRACK_AS) = "As, " & OutputAreaUnit(units) & " (растянутая арматура)"
    data(COLUMN_HEADER_ROW, COL_CRACK_ABT) = "Abt, " & OutputAreaUnit(units) & " (растянутый бетон)"
    data(COLUMN_HEADER_ROW, COL_CRACK_DS) = "ds_eq, " & OutputLengthUnit(units) & " (экв. диаметр)"
    data(COLUMN_HEADER_ROW, COL_CRACK_REBAR_COUNT) = "TensionRebarCount (стержней в зоне)"
    data(COLUMN_HEADER_ROW, COL_CRACK_LS_RAW) = "ls_raw, " & OutputLengthUnit(units) & " (шаг до ограничений)"
    data(COLUMN_HEADER_ROW, COL_CRACK_LS) = "ls, " & OutputLengthUnit(units) & " (принятый шаг)"
    data(COLUMN_HEADER_ROW, COL_CRACK_PHI1) = "phi1 (длительность)"
    data(COLUMN_HEADER_ROW, COL_CRACK_PHI2) = "phi2 (коэффициент из Config)"
    data(COLUMN_HEADER_ROW, COL_CRACK_PHI3) = "phi3 (характер нагрузки)"
    data(COLUMN_HEADER_ROW, COL_CRACK_NCRC) = "Ncrc, " & OutputForceUnit(units) & " (усилие образования трещин при центральном растяжении)"
    data(COLUMN_HEADER_ROW, COL_CRACK_LAMBDA_CRC) = "lambda_crc (множитель трещины)"
    data(COLUMN_HEADER_ROW, COL_CRACK_SIGMA_S) = "sigma_s, " & OutputStressUnit(units) & " (напряжение арматуры)"
    data(COLUMN_HEADER_ROW, COL_CRACK_SIGMA_S_CRC) = "sigma_s,crc, " & OutputStressUnit(units) & " (при образовании трещины)"
    data(COLUMN_HEADER_ROW, COL_CRACK_PSI) = "psi_s (работа растянутого бетона)"
    data(COLUMN_HEADER_ROW, COL_CRACK_ES) = "Es, " & OutputStressUnit(units) & " (модуль арматуры)"
    data(COLUMN_HEADER_ROW, COL_CRACK_WIDTH) = "a_crc, " & OutputLengthUnit(units) & " (ширина трещины)"
    data(COLUMN_HEADER_ROW, COL_CRACK_ALLOWABLE) = "a_crc,ult, " & OutputLengthUnit(units) & " (допустимая ширина)"
    data(COLUMN_HEADER_ROW, COL_CRACK_SF) = "CrackSafetyFactor (запас по трещинам)"

    data(COLUMN_HEADER_ROW, COL_LONG_CRACK_STATUS) = "LongitudinalCrackStatus (продольные трещины)"
    data(COLUMN_HEADER_ROW, COL_LONG_CRACK_SIGMA_C_MAX) = "sigma_c,max, " & OutputStressUnit(units) & " (сжатие бетона)"
    data(COLUMN_HEADER_ROW, COL_LONG_CRACK_RB_MC2) = "Rb,mc2, " & OutputStressUnit(units) & " (допустимое сжатие)"
    data(COLUMN_HEADER_ROW, COL_LONG_CRACK_SF) = "LongitudinalCrackSafetyFactor (запас)"

    data(COLUMN_HEADER_ROW, COL_STABILITY_STATUS) = "Статус устойчивости"
    data(COLUMN_HEADER_ROW, COL_STABILITY_NCR_MIN) = "Контрольная сжимающая N, " & OutputForceUnit(units)
    data(COLUMN_HEADER_ROW, COL_STABILITY_RESERVE) = "Запас устойчивости"
    data(COLUMN_HEADER_ROW, COL_STABILITY_DESIGN_N) = "N после устойчивости, " & OutputForceUnit(units)
    data(COLUMN_HEADER_ROW, COL_STABILITY_LOADPOINT_MX) = "Mx после устойчивости в точке нагрузки, " & OutputMomentUnit(units)
    data(COLUMN_HEADER_ROW, COL_STABILITY_LOADPOINT_MY) = "My после устойчивости в точке нагрузки, " & OutputMomentUnit(units)
    data(COLUMN_HEADER_ROW, COL_STABILITY_DESIGN_MX) = "Mx после устойчивости в ц.т. приведенного сечения, " & OutputMomentUnit(units)
    data(COLUMN_HEADER_ROW, COL_STABILITY_DESIGN_MY) = "My после устойчивости в ц.т. приведенного сечения, " & OutputMomentUnit(units)
    PutStabilityPlaneHeaders data, COL_STABILITY_SP35_NCR1, units, "SP35"
    PutStabilityPlaneHeaders data, COL_STABILITY_SP35_NCR2, units, "SP35"
    PutStabilityPlaneHeaders data, COL_STABILITY_SP63_NCR1, units, "SP63"
    PutStabilityPlaneHeaders data, COL_STABILITY_SP63_NCR2, units, "SP63"

    data(COLUMN_HEADER_ROW, COL_OVERALL_SF) = "MinSafetyFactor (общий запас)"
End Sub

' Заполняет набор заголовков для одной расчетной плоскости устойчивости.
' Математически поля одинаковые, но обозначения эксцентриситетов в СП 35
' и СП 63 разные, поэтому подписи формируются под конкретный нормативный блок.
Private Sub PutStabilityPlaneHeaders(ByRef data() As Variant, ByVal firstColumn As Long, _
        ByVal units As CUnitSystem, ByVal codeName As String)
    data(COLUMN_HEADER_ROW, firstColumn) = "Ncr, " & OutputForceUnit(units)
    data(COLUMN_HEADER_ROW, firstColumn + 1) = "Nult,stab, " & OutputForceUnit(units)
    data(COLUMN_HEADER_ROW, firstColumn + 2) = "Запас"
    data(COLUMN_HEADER_ROW, firstColumn + 3) = "eta"
    data(COLUMN_HEADER_ROW, firstColumn + 4) = "phi_l"
    If StrComp(codeName, "SP35", vbTextCompare) = 0 Then
        data(COLUMN_HEADER_ROW, firstColumn + 5) = "Расчетный ec, " & OutputLengthUnit(units)
        data(COLUMN_HEADER_ROW, firstColumn + 6) = "Случайный ec,sl, " & OutputLengthUnit(units)
    Else
        data(COLUMN_HEADER_ROW, firstColumn + 5) = "Начальный e0, " & OutputLengthUnit(units)
        data(COLUMN_HEADER_ROW, firstColumn + 6) = "Случайный ea, " & OutputLengthUnit(units)
    End If
    data(COLUMN_HEADER_ROW, firstColumn + 7) = "Расчетная длина l0, " & OutputLengthUnit(units)
    data(COLUMN_HEADER_ROW, firstColumn + 8) = "Жесткость D, " & OutputStiffnessUnit(units)
End Sub

Private Sub PutCombination(ByRef data() As Variant, ByVal rowIndex As Long, ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal units As CUnitSystem, ByVal section As CSectionModel)
    Dim policy As CBatchStatusPolicy
    Set policy = StatusPolicy()

    If rowIndex > UBound(data, 1) Then Exit Sub

    data(rowIndex, COL_OVERALL_STATUS) = policy.DisplayStatus(batch.OverallStatus(index))
    data(rowIndex, COL_ID) = batch.CombinationID(index)
    data(rowIndex, COL_COMMENT) = batch.CombinationName(index)
    data(rowIndex, COL_PROFILE_ID) = batch.ProfileId(index)
    data(rowIndex, COL_DIRECT_STATUS) = policy.DisplayStatus(batch.DirectStateStatus(index))

    If batch.StateAvailable(index) Then
        data(rowIndex, COL_EPS_C_COMP) = batch.MinConcreteStrain(index)
        data(rowIndex, COL_EPS_C_TENS) = batch.MaxConcreteStrain(index)
        data(rowIndex, COL_EPS_S_COMP) = batch.MinSteelStrain(index)
        data(rowIndex, COL_EPS_S_TENS) = batch.MaxSteelStrain(index)
        PutCurrentStateGeometry data, rowIndex, batch, index, units, section
        data(rowIndex, COL_LIM_C_COMP) = batch.ConcreteCompressionLimitForCombination(index)
        If batch.ConcreteTensionLimitEnabledForCombination(index) Then data(rowIndex, COL_LIM_C_TENS) = batch.ConcreteTensionLimitForCombination(index)
        data(rowIndex, COL_LIM_S_COMP) = batch.SteelCompressionLimitForCombination(index)
        data(rowIndex, COL_LIM_S_TENS) = batch.SteelTensionLimitForCombination(index)
    End If

    data(rowIndex, COL_CAPACITY_STATUS) = policy.DisplayStatus(batch.CapacityStatus(index))
    data(rowIndex, COL_CAPACITY_LIMIT_STATE) = policy.DisplayStatus(batch.CapacityLimitState(index))
    If policy.DisplayStatus(batch.CapacityStatus(index)) <> policy.NotApplicable Then
        data(rowIndex, COL_CAPACITY_LOAD_PATH) = batch.CapacityLoadPath(index)
    End If
    If policy.IsFinished(batch.CapacityStatus(index)) Then
        data(rowIndex, COL_CAPACITY_SOLUTION_METHOD) = batch.CapacitySolutionMethod(index)
        data(rowIndex, COL_CAPACITY_LAMBDA) = batch.LambdaCapacity(index)
        If CapacityLoadPathScalesN(batch.CapacityLoadPathKey(index)) Then data(rowIndex, COL_CAPACITY_N_ULT) = OutputForce(batch.NUltimate(index), units)
        If CapacityLoadPathScalesMx(batch.CapacityLoadPathKey(index)) Then data(rowIndex, COL_CAPACITY_MX_ULT) = OutputMx(batch.MxUltimate(index), units)
        If CapacityLoadPathScalesMy(batch.CapacityLoadPathKey(index)) Then data(rowIndex, COL_CAPACITY_MY_ULT) = OutputMy(batch.MyUltimate(index), units)
        data(rowIndex, COL_CAPACITY_SF) = CapacitySafetyFormula(rowIndex)
        PutCapacityStateGeometry data, rowIndex, batch, index, units, section
    End If

    data(rowIndex, COL_CRACK_STATUS) = policy.DisplayStatus(batch.CrackStatus(index))
    data(rowIndex, COL_CRACK_NCRC) = NOT_CALCULATED_SHORT
    If policy.IsFinished(batch.CrackStatus(index)) Then
        data(rowIndex, COL_CRACK_FORMED) = IIf(batch.CrackFormed(index), "Да", "Нет")
        data(rowIndex, COL_CRACK_H) = OutputLength(batch.CrackSectionDepthH(index), units)
        data(rowIndex, COL_CRACK_A) = OutputLength(batch.CrackCoverA(index), units)
        data(rowIndex, COL_CRACK_TENSION_DEPTH) = OutputLength(batch.CrackTensionDepth(index), units)
        data(rowIndex, COL_CRACK_EFFECTIVE_DEPTH) = OutputLength(batch.CrackEffectiveZoneDepth(index), units)
        data(rowIndex, COL_CRACK_AS) = OutputArea(batch.CrackAs(index), units)
        data(rowIndex, COL_CRACK_ABT) = OutputArea(batch.CrackAbt(index), units)
        data(rowIndex, COL_CRACK_DS) = OutputLength(batch.CrackDsEquivalent(index), units)
        data(rowIndex, COL_CRACK_REBAR_COUNT) = batch.CrackTensionRebarCount(index)
        data(rowIndex, COL_CRACK_LS_RAW) = OutputLength(batch.CrackSpacingRaw(index), units)
        data(rowIndex, COL_CRACK_LS) = OutputLength(batch.CrackSpacing(index), units)
        data(rowIndex, COL_CRACK_PHI1) = batch.CrackPhi1(index)
        data(rowIndex, COL_CRACK_PHI2) = batch.CrackPhi2(index)
        data(rowIndex, COL_CRACK_PHI3) = batch.CrackPhi3(index)
        If batch.CrackNcrc(index) > 0# Then
            data(rowIndex, COL_CRACK_NCRC) = OutputForce(batch.CrackNcrc(index), units)
        Else
            data(rowIndex, COL_CRACK_NCRC) = NOT_CALCULATED_SHORT
        End If
        If batch.CrackSigmaSCrc(index) > 0# Then
            data(rowIndex, COL_CRACK_LAMBDA_CRC) = batch.CrackLambdaCrc(index)
        Else
            data(rowIndex, COL_CRACK_LAMBDA_CRC) = NOT_CALCULATED_SHORT
        End If
        data(rowIndex, COL_CRACK_SIGMA_S) = OutputStress(batch.CrackSigmaS(index), units)
        If batch.CrackSigmaSCrc(index) > 0# Then
            data(rowIndex, COL_CRACK_SIGMA_S_CRC) = OutputStress(batch.CrackSigmaSCrc(index), units)
        Else
            data(rowIndex, COL_CRACK_SIGMA_S_CRC) = NOT_CALCULATED_SHORT
        End If
        data(rowIndex, COL_CRACK_PSI) = batch.CrackPsiS(index)
        data(rowIndex, COL_CRACK_ES) = OutputStress(batch.CrackSteelElasticModulus, units)
        data(rowIndex, COL_CRACK_WIDTH) = CrackWidthFormula(rowIndex)
        data(rowIndex, COL_CRACK_ALLOWABLE) = OutputLength(batch.CrackAllowable(index), units)
        data(rowIndex, COL_CRACK_SF) = CrackSafetyFormula(rowIndex)
    End If

    data(rowIndex, COL_LONG_CRACK_STATUS) = policy.DisplayStatus(batch.LongitudinalCrackStatus(index))
    If policy.IsFinished(batch.LongitudinalCrackStatus(index)) Then
        data(rowIndex, COL_LONG_CRACK_SIGMA_C_MAX) = OutputStress(batch.MaxConcreteCompressionStress(index), units)
        data(rowIndex, COL_LONG_CRACK_RB_MC2) = OutputStress(batch.LongitudinalCrackRbMc2(index), units)
        data(rowIndex, COL_LONG_CRACK_SF) = LongitudinalCrackSafetyFormula(rowIndex)
    End If

    data(rowIndex, COL_STABILITY_STATUS) = policy.DisplayStatus(batch.StabilityStatus(index))
    If policy.DisplayStatus(batch.StabilityStatus(index)) <> policy.NotApplicable Then
        data(rowIndex, COL_STABILITY_NCR_MIN) = OutputCompressionForce(batch.StabilityCriticalForce(index), units)
        data(rowIndex, COL_STABILITY_RESERVE) = StabilitySafetyFormula(rowIndex)
        data(rowIndex, COL_STABILITY_DESIGN_N) = OutputForce(batch.StabilityDesignN(index), units)
        data(rowIndex, COL_STABILITY_LOADPOINT_MX) = OutputMx(batch.StabilityLoadPointMx(index), units)
        data(rowIndex, COL_STABILITY_LOADPOINT_MY) = OutputMy(batch.StabilityLoadPointMy(index), units)
        data(rowIndex, COL_STABILITY_DESIGN_MX) = OutputMx(batch.StabilityDesignMx(index), units)
        data(rowIndex, COL_STABILITY_DESIGN_MY) = OutputMy(batch.StabilityDesignMy(index), units)
        PutStabilityDetailsForSelectedCode data, rowIndex, batch, index, units
    End If

    data(rowIndex, COL_OVERALL_SF) = OverallSafetyFormula(rowIndex)
End Sub

' Заполняет только тот нормативный блок устойчивости, который реально выбран
' настройкой Stability.Code. Соседний блок остается пустым, чтобы данные СП 35
' и СП 63 не смешивались в одной группе колонок.
Private Sub PutStabilityDetailsForSelectedCode(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal units As CUnitSystem)
    Dim codeText As String
    codeText = Trim$(batch.StabilityCode(index))
    If StrComp(codeText, "SP35", vbTextCompare) = 0 Then
        PutStabilityPlaneValues data, rowIndex, COL_STABILITY_SP35_NCR1, batch, index, 1, units
        PutStabilityPlaneValues data, rowIndex, COL_STABILITY_SP35_NCR2, batch, index, 2, units
    ElseIf StrComp(codeText, "SP63", vbTextCompare) = 0 Then
        PutStabilityPlaneValues data, rowIndex, COL_STABILITY_SP63_NCR1, batch, index, 1, units
        PutStabilityPlaneValues data, rowIndex, COL_STABILITY_SP63_NCR2, batch, index, 2, units
    End If
End Sub

' Записывает данные одной главной плоскости устойчивости в выбранный блок СП.
Private Sub PutStabilityPlaneValues(ByRef data() As Variant, ByVal rowIndex As Long, ByVal firstColumn As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long, ByVal units As CUnitSystem)
    Dim notApplicableText As String
    notApplicableText = StatusPolicy().NotApplicable

    If planeIndex = 1 Then
        data(rowIndex, firstColumn) = OptionalCompressionForce(batch.StabilityNcr1(index), units, notApplicableText)
        data(rowIndex, firstColumn + 1) = OptionalCompressionForce(batch.StabilityNultimate1(index), units, notApplicableText)
        data(rowIndex, firstColumn + 2) = batch.StabilityReserve1(index)
        data(rowIndex, firstColumn + 3) = batch.StabilityEta1(index)
        data(rowIndex, firstColumn + 4) = batch.StabilityPhiL1(index)
        data(rowIndex, firstColumn + 5) = OutputLength(batch.StabilityEccentricity1(index), units)
        data(rowIndex, firstColumn + 6) = OutputLength(batch.StabilityAccidentalEcc1(index), units)
        data(rowIndex, firstColumn + 7) = OutputLength(batch.StabilityEffectiveLength1(index), units)
        data(rowIndex, firstColumn + 8) = OutputStiffness(batch.StabilityStiffnessD1(index), units)
    Else
        data(rowIndex, firstColumn) = OptionalCompressionForce(batch.StabilityNcr2(index), units, notApplicableText)
        data(rowIndex, firstColumn + 1) = OptionalCompressionForce(batch.StabilityNultimate2(index), units, notApplicableText)
        data(rowIndex, firstColumn + 2) = batch.StabilityReserve2(index)
        data(rowIndex, firstColumn + 3) = batch.StabilityEta2(index)
        data(rowIndex, firstColumn + 4) = batch.StabilityPhiL2(index)
        data(rowIndex, firstColumn + 5) = OutputLength(batch.StabilityEccentricity2(index), units)
        data(rowIndex, firstColumn + 6) = OutputLength(batch.StabilityAccidentalEcc2(index), units)
        data(rowIndex, firstColumn + 7) = OutputLength(batch.StabilityEffectiveLength2(index), units)
        data(rowIndex, firstColumn + 8) = OutputStiffness(batch.StabilityStiffnessD2(index), units)
    End If
End Sub

' Записывает h/x для фактического состояния LC. Эти величины помогают
' интерпретировать текущие деформации и не являются параметрами предельного
' момента; поэтому они находятся в блоке проверки по относительным деформациям.
Private Sub PutCurrentStateGeometry(ByRef data() As Variant, ByVal rowIndex As Long, ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal units As CUnitSystem, ByVal section As CSectionModel)
    If section Is Nothing Then Exit Sub
    If Not batch.StateAvailable(index) Then Exit Sub

    Dim depthH As Double
    Dim compressedDepth As Double
    CalculateDepthsPerpendicularToNeutral section, batch.Epsilon0(index), batch.KappaX(index), batch.KappaY(index), _
        depthH, compressedDepth
    If depthH > 0# Then data(rowIndex, COL_STATE_H) = OutputLength(depthH, units)
    If compressedDepth >= 0# Then data(rowIndex, COL_STATE_C) = OutputLength(compressedDepth, units)
End Sub

' Записывает h/x только для состояния найденного предельного момента.
' В режиме DirectState или при NumericalFailure этих данных нет, поэтому
' ячейки остаются пустыми и не смешиваются с фактическим состоянием LC.
' Записывает h/x для предельного состояния, найденного на выбранной lambda-траектории.
' Эти величины относятся не к текущему LC, а именно к состоянию lambdaUltimate.
Private Sub PutCapacityStateGeometry(ByRef data() As Variant, ByVal rowIndex As Long, ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal units As CUnitSystem, ByVal section As CSectionModel)
    If section Is Nothing Then Exit Sub
    If Not batch.CapacityStateAvailable(index) Then Exit Sub

    Dim depthH As Double
    Dim compressedDepth As Double
    CalculateDepthsPerpendicularToNeutral section, batch.CapacityEpsilon0(index), batch.CapacityKappaX(index), batch.CapacityKappaY(index), _
        depthH, compressedDepth
    If depthH > 0# Then data(rowIndex, COL_CAPACITY_H) = OutputLength(depthH, units)
    If compressedDepth >= 0# Then data(rowIndex, COL_CAPACITY_C) = OutputLength(compressedDepth, units)
End Sub

' Определяет габарит h и глубину сжатой зоны c вдоль нормали к нейтральной линии.
' Сам габарит берет из CSectionPropertiesCalculator, чтобы Summary, устойчивость
' и другие проверки пользовались одной логикой проекций расчетной геометрии.
Private Sub CalculateDepthsPerpendicularToNeutral(ByVal section As CSectionModel, ByVal eps0 As Double, _
        ByVal kappaX As Double, ByVal kappaY As Double, ByRef depthH As Double, ByRef compressedDepth As Double)
    Dim normValue As Double
    normValue = Sqr(kappaX * kappaX + kappaY * kappaY)
    If normValue <= 0.000000000000001 Then Exit Sub

    Dim nx As Double
    Dim ny As Double
    nx = kappaY / normValue
    ny = kappaX / normValue

    Dim minS As Double
    Dim maxS As Double
    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateProjection section, nx, ny, True, minS, maxS

    depthH = maxS - minS
    If depthH <= 0# Then Exit Sub

    Dim neutralS As Double
    neutralS = -eps0 / normValue

    If eps0 + normValue * minS < 0# And eps0 + normValue * maxS < 0# Then
        compressedDepth = depthH
    ElseIf eps0 + normValue * minS >= 0# And eps0 + normValue * maxS >= 0# Then
        compressedDepth = 0#
    Else
        compressedDepth = MaxDouble(0#, MinDouble(maxS, neutralS) - minS)
    End If
End Sub

Private Function CapacitySafetyFormula(ByVal rowIndex As Long) As String
    CapacitySafetyFormula = "=IFERROR(" & CellAddress(rowIndex, COL_CAPACITY_LAMBDA) & ","""")"
End Function

Private Function CrackWidthFormula(ByVal rowIndex As Long) As String
    CrackWidthFormula = "=IF(" & CellAddress(rowIndex, COL_CRACK_FORMED) & "=""" & "Да" & """," & _
        CellAddress(rowIndex, COL_CRACK_PHI1) & "*" & CellAddress(rowIndex, COL_CRACK_PHI2) & "*" & _
        CellAddress(rowIndex, COL_CRACK_PHI3) & "*" & CellAddress(rowIndex, COL_CRACK_PSI) & "*(" & _
        CellAddress(rowIndex, COL_CRACK_SIGMA_S) & "/" & CellAddress(rowIndex, COL_CRACK_ES) & ")*" & _
        CellAddress(rowIndex, COL_CRACK_LS) & ",0)"
End Function

Private Function CrackSafetyFormula(ByVal rowIndex As Long) As String
    CrackSafetyFormula = "=IFERROR(IF(" & CellAddress(rowIndex, COL_CRACK_WIDTH) & ">0," & _
        CellAddress(rowIndex, COL_CRACK_ALLOWABLE) & "/" & CellAddress(rowIndex, COL_CRACK_WIDTH) & ",""""),"""")"
End Function

Private Function LongitudinalCrackSafetyFormula(ByVal rowIndex As Long) As String
    LongitudinalCrackSafetyFormula = "=IFERROR(IF(" & CellAddress(rowIndex, COL_LONG_CRACK_SIGMA_C_MAX) & ">0," & _
        CellAddress(rowIndex, COL_LONG_CRACK_RB_MC2) & "/" & _
        CellAddress(rowIndex, COL_LONG_CRACK_SIGMA_C_MAX) & ",""""),"""")"
End Function

Private Function StabilitySafetyFormula(ByVal rowIndex As Long) As String
    StabilitySafetyFormula = "=IFERROR(IF(ABS(" & CellAddress(rowIndex, COL_STABILITY_NCR_MIN) & ")>0," & _
        "ABS(" & CellAddress(rowIndex, COL_STABILITY_NCR_MIN) & ")/" & _
        "ABS(" & CellAddress(rowIndex, COL_STABILITY_DESIGN_N) & "),""""),"""")"
End Function

Private Function OverallSafetyFormula(ByVal rowIndex As Long) As String
    OverallSafetyFormula = "=IFERROR(AGGREGATE(15,6,CHOOSE({1,2,3,4}," & _
        "IF(" & CellAddress(rowIndex, COL_CAPACITY_SF) & ">0," & CellAddress(rowIndex, COL_CAPACITY_SF) & ",NA())," & _
        "IF(" & CellAddress(rowIndex, COL_CRACK_SF) & ">0," & CellAddress(rowIndex, COL_CRACK_SF) & ",NA())," & _
        "IF(" & CellAddress(rowIndex, COL_LONG_CRACK_SF) & ">0," & CellAddress(rowIndex, COL_LONG_CRACK_SF) & ",NA())," & _
        "IF(" & CellAddress(rowIndex, COL_STABILITY_RESERVE) & ">0," & CellAddress(rowIndex, COL_STABILITY_RESERVE) & ",NA())),1),"""")"
End Function

Private Function CellAddress(ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    CellAddress = ColumnLetter(columnIndex) & CStr(rowIndex)
End Function

Private Function ColumnLetter(ByVal columnIndex As Long) As String
    Dim dividend As Long
    Dim moduloValue As Long
    dividend = columnIndex
    Do While dividend > 0
        moduloValue = (dividend - 1) Mod 26
        ColumnLetter = Chr$(65 + moduloValue) & ColumnLetter
        dividend = (dividend - moduloValue) \ 26
    Loop
End Function

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

    stage = "HeaderStyle"
    target.Rows.Item(1).Font.Bold = True
    target.Rows.Item(GROUP_HEADER_ROW).Font.Bold = True
    target.Rows.Item(SUBGROUP_HEADER_ROW).Font.Bold = True
    target.Rows.Item(COLUMN_HEADER_ROW).Font.Bold = True
    target.Rows.Item(GROUP_HEADER_ROW).Interior.Color = HEADER_COLOR
    target.Rows.Item(SUBGROUP_HEADER_ROW).Interior.Color = HEADER_COLOR
    target.Rows.Item(COLUMN_HEADER_ROW).Interior.Color = HEADER_COLOR
    target.Rows.Item(1).Interior.Color = HEADER_COLOR

    stage = "MergeHeaders"
    MergeSummaryHeader target, GROUP_HEADER_ROW, COL_DIRECT_STATUS, COL_LIM_S_TENS
    MergeSummaryHeader target, GROUP_HEADER_ROW, COL_CAPACITY_STATUS, COL_CAPACITY_SF
    MergeSummaryHeader target, GROUP_HEADER_ROW, COL_CRACK_STATUS, COL_CRACK_SF
    MergeSummaryHeader target, GROUP_HEADER_ROW, COL_LONG_CRACK_STATUS, COL_LONG_CRACK_SF
    MergeSummaryHeader target, GROUP_HEADER_ROW, COL_STABILITY_STATUS, COL_STABILITY_SP63_D2
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_OVERALL_STATUS, COL_PROFILE_ID
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_DIRECT_STATUS, COL_STATE_C
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_LIM_C_COMP, COL_LIM_S_TENS
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_CAPACITY_STATUS, COL_CAPACITY_SOLUTION_METHOD
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_CAPACITY_N_ULT, COL_CAPACITY_MY_ULT
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_CAPACITY_H, COL_CAPACITY_C
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_CRACK_STATUS, COL_CRACK_EFFECTIVE_DEPTH
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_CRACK_AS, COL_CRACK_REBAR_COUNT
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_CRACK_LS_RAW, COL_CRACK_LS
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_CRACK_PHI1, COL_CRACK_PHI3
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_CRACK_NCRC, COL_CRACK_PSI
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_CRACK_WIDTH, COL_CRACK_SF
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_LONG_CRACK_STATUS, COL_LONG_CRACK_SF
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_STABILITY_STATUS, COL_STABILITY_DESIGN_MY
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_STABILITY_SP35_NCR1, COL_STABILITY_SP35_D1
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_STABILITY_SP35_NCR2, COL_STABILITY_SP35_D2
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_STABILITY_SP63_NCR1, COL_STABILITY_SP63_D1
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_STABILITY_SP63_NCR2, COL_STABILITY_SP63_D2

    stage = "Alignment"
    target.Rows.Item(COLUMN_HEADER_ROW).HorizontalAlignment = -4131
    target.Rows.Item(GROUP_HEADER_ROW).HorizontalAlignment = -4108
    target.Rows.Item(SUBGROUP_HEADER_ROW).HorizontalAlignment = -4108
    target.Rows.Item(1).HorizontalAlignment = -4131
    target.Offset(0, 0).Resize(META_ROWS, TARGET_COLUMNS).HorizontalAlignment = -4131

    Dim dataRows As Long
    dataRows = TARGET_ROWS - FIRST_DATA_ROW + 1
    target.Offset(FIRST_DATA_ROW - 1, 0).Resize(dataRows, TARGET_COLUMNS).HorizontalAlignment = -4108
    target.Offset(FIRST_DATA_ROW - 1, COL_COMMENT - 1).Resize(dataRows, 1).HorizontalAlignment = -4152
    target.Offset(FIRST_DATA_ROW - 1, COL_CAPACITY_SF - 1).Resize(dataRows, 1).Interior.Color = SAFETY_COLOR
    target.Offset(FIRST_DATA_ROW - 1, COL_CRACK_SF - 1).Resize(dataRows, 1).Interior.Color = SAFETY_COLOR
    target.Offset(FIRST_DATA_ROW - 1, COL_LONG_CRACK_SF - 1).Resize(dataRows, 1).Interior.Color = SAFETY_COLOR
    target.Offset(FIRST_DATA_ROW - 1, COL_STABILITY_RESERVE - 1).Resize(dataRows, 1).Interior.Color = SAFETY_COLOR
    target.Offset(FIRST_DATA_ROW - 1, COL_STABILITY_SP35_RESERVE1 - 1).Resize(dataRows, 1).Interior.Color = SAFETY_COLOR
    target.Offset(FIRST_DATA_ROW - 1, COL_STABILITY_SP35_RESERVE2 - 1).Resize(dataRows, 1).Interior.Color = SAFETY_COLOR
    target.Offset(FIRST_DATA_ROW - 1, COL_STABILITY_SP63_RESERVE1 - 1).Resize(dataRows, 1).Interior.Color = SAFETY_COLOR
    target.Offset(FIRST_DATA_ROW - 1, COL_STABILITY_SP63_RESERVE2 - 1).Resize(dataRows, 1).Interior.Color = SAFETY_COLOR
    target.Offset(FIRST_DATA_ROW - 1, COL_OVERALL_SF - 1).Resize(dataRows, 1).Interior.Color = OVERALL_SAFETY_COLOR
    stage = "FormatOverallStatusRows"
    FormatOverallStatusRows target, batch

    stage = "Borders"
    Dim borderedArea As Object
    Set borderedArea = target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(TARGET_ROWS - GROUP_HEADER_ROW + 1, TARGET_COLUMNS)
    borderedArea.Borders.LineStyle = 1
    borderedArea.Borders.Color = RGB(0, 0, 0)
    borderedArea.Borders.Weight = 2
    stage = "NumberFormat"
    target.Offset(0, 0).Resize(1, TARGET_COLUMNS).NumberFormat = "@"
    stage = "WriteStatusDictionary"
    WriteStatusDictionary target
    Exit Sub

Failed:
    Err.Raise Err.Number, "CBatchResultWriter.FormatSummary", stage & ": " & Err.Description
End Sub

Private Sub MergeSummaryHeader(ByVal target As Object, ByVal rowIndex As Long, ByVal firstColumn As Long, ByVal lastColumn As Long)
    Dim area As Object
    Set area = target.Offset(rowIndex - 1, firstColumn - 1).Resize(1, lastColumn - firstColumn + 1)
    area.Merge
End Sub

' Красит всю строку LC, если общий статус показывает физический отказ,
' численную несходимость или ошибку исходных данных.
Private Sub FormatOverallStatusRows(ByVal target As Object, ByVal batch As CBatchSectionCalculator)
    If batch Is Nothing Then Exit Sub
    If batch.Count <= 0 Then Exit Sub

    Dim policy As CBatchStatusPolicy
    Set policy = StatusPolicy()

    Dim i As Long
    For i = 1 To batch.Count
        Dim statusText As String
        statusText = batch.OverallStatus(i)
        If policy.IsFailedOverall(statusText) Then
            With target.Offset(FIRST_DATA_ROW + i - 2, 0).Resize(1, TARGET_COLUMNS)
                .Interior.Color = FAILED_ROW_COLOR
                .Font.Color = FAILED_FONT_COLOR
            End With
        End If
    Next i
End Sub

' Выводит справа от Summary компактный справочник пользовательских статусов.
' Таблица не входит в rngBatchSummary, но каждый расчет обновляет ее вместе
' со сводкой, чтобы пользователь видел расшифровку ровно тех кодов, которые
' могут появиться в строках сочетаний.
Private Sub WriteStatusDictionary(ByVal target As Object)
    Dim topLeft As Object
    Dim dictionaryColumn As Long
    Dim policy As CBatchStatusPolicy
    Set policy = StatusPolicy()

    dictionaryColumn = target.Column + TARGET_COLUMNS + 1

    ' Словарь статусов стоит справа от Summary, но начинается ниже мета-строк.
    ' Так длинные расшифровки не раздувают первые строки с итогами расчета.
    Set topLeft = target.Worksheet.Cells(target.Row, dictionaryColumn)
    topLeft.Resize(GROUP_HEADER_ROW + 6, 2).Clear
    topLeft.Resize(GROUP_HEADER_ROW + 6, 2).WrapText = False
    Dim resetRow As Long
    For resetRow = target.Row To target.Row + GROUP_HEADER_ROW - 2
        target.Worksheet.Rows.Item(resetRow).RowHeight = 15
    Next resetRow
    Set topLeft = target.Worksheet.Cells(target.Row + GROUP_HEADER_ROW - 1, dictionaryColumn)

    topLeft.Value2 = "Статус"
    topLeft.Offset(0, 1).Value2 = "Расшифровка"
    topLeft.Offset(1, 0).Value2 = policy.OK
    topLeft.Offset(1, 1).Value2 = "Расчет этой проверки выполнен успешно. Для прямого НДС равновесие найдено внутри физической диаграммы материала; для прочности или трещин проверка прошла."
    topLeft.Offset(2, 0).Value2 = policy.Fail
    topLeft.Offset(2, 1).Value2 = "Численное решение есть, но инженерная проверка не проходит: превышена несущая способность, ширина нормальной трещины больше допуска, sigma_c,max больше Rb,mc2 по продольным трещинам или прямое НДС найдено только за пределами физической диаграммы."
    topLeft.Offset(3, 0).Value2 = policy.NumFail
    topLeft.Offset(3, 1).Value2 = "Solver не смог найти равновесие или предельную точку с текущими численными настройками. Это не доказательство разрушения; нужно смотреть диагностику, нагрузку, сетку и параметры solver-а."
    topLeft.Offset(4, 0).Value2 = policy.InputErr
    topLeft.Offset(4, 1).Value2 = "Расчет не запускался из-за ошибки во входных данных: пустая или неизвестная настройка, недопустимая геометрия, материал, единицы или строка сочетания."
    topLeft.Offset(5, 0).Value2 = policy.NotApplicable
    topLeft.Offset(5, 1).Value2 = "Проверка для этого сочетания не выполнялась и не считается ошибкой: например, она не включена выбранным профилем или неприменима к найденному состоянию."

    With topLeft.Resize(1, 2)
        .Font.Bold = True
        .Interior.Color = HEADER_COLOR
    End With
    With topLeft.Resize(6, 2)
        .Borders.LineStyle = 1
        .Borders.Color = RGB(0, 0, 0)
        .HorizontalAlignment = -4131
        .VerticalAlignment = -4108
        .WrapText = False
    End With
End Sub

Private Function StatusPolicy() As CBatchStatusPolicy
    Static policy As CBatchStatusPolicy
    If policy Is Nothing Then Set policy = New CBatchStatusPolicy
    Set StatusPolicy = policy
End Function

Private Function OutputLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputLength = value Else OutputLength = units.InternalLengthToOutput(value)
End Function

Private Function OutputLengthForSummary(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If Abs(value) < 0.000000001 Then value = 0#
    OutputLengthForSummary = OutputLength(value, units)
End Function

Private Function OutputLengthUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputLengthUnit = "mm" Else OutputLengthUnit = units.OutputLengthUnit
End Function

Private Function OutputArea(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputArea = value Else OutputArea = units.InternalAreaToOutput(value)
End Function

Private Function OutputAreaUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputAreaUnit = "mm2" Else OutputAreaUnit = units.OutputAreaUnit
End Function

Private Function OutputStress(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputStress = value Else OutputStress = units.InternalStressToOutput(value)
End Function

Private Function OutputStressUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputStressUnit = "MPa" Else OutputStressUnit = units.OutputStressUnit
End Function

Private Function OutputForce(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputForce = value Else OutputForce = units.InternalForceToOutput(value)
End Function

' Выводит критическую сжимающую силу со знаком пользовательской системы N.
' Внутри расчета Ncr хранится положительной величиной сжатия, поэтому перед
' обратным преобразованием ей возвращается внутренний знак сжатия.
Private Function OutputCompressionForce(ByVal compressionMagnitude As Double, ByVal units As CUnitSystem) As Double
    If compressionMagnitude <= 0# Then Exit Function
    OutputCompressionForce = OutputForce(-Abs(compressionMagnitude), units)
End Function

Private Function OptionalCompressionForce(ByVal compressionMagnitude As Double, _
        ByVal units As CUnitSystem, ByVal emptyText As String) As Variant
    If compressionMagnitude > 0# Then
        OptionalCompressionForce = OutputCompressionForce(compressionMagnitude, units)
    Else
        OptionalCompressionForce = emptyText
    End If
End Function

Private Function OutputForceUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputForceUnit = "N" Else OutputForceUnit = units.OutputForceUnit
End Function

Private Function OutputStiffness(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputStiffness = value Else OutputStiffness = units.InternalForceLengthSquaredToOutput(value)
End Function

Private Function OutputStiffnessUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputStiffnessUnit = "N*mm2" Else OutputStiffnessUnit = units.OutputForceLengthSquaredUnit
End Function

Private Function OutputMx(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMx = value Else OutputMx = units.InternalMomentMxToOutput(value)
End Function

Private Function OutputMy(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMy = value Else OutputMy = units.InternalMomentMyToOutput(value)
End Function

Private Function OutputMomentMagnitude(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMomentMagnitude = value Else OutputMomentMagnitude = units.InternalMomentMagnitudeToOutput(value)
End Function

Private Function OutputMomentUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputMomentUnit = "N*mm" Else OutputMomentUnit = units.OutputMomentUnit
End Function

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

Private Function CapacityLoadPathScalesN(ByVal loadPath As String) As Boolean
    CapacityLoadPathScalesN = (StrComp(loadPath, "LambdaN", vbTextCompare) = 0 Or _
        StrComp(loadPath, "LambdaNMxy", vbTextCompare) = 0)
End Function

Private Function CapacityLoadPathScalesMx(ByVal loadPath As String) As Boolean
    CapacityLoadPathScalesMx = (StrComp(loadPath, "LambdaMx", vbTextCompare) = 0 Or _
        StrComp(loadPath, "LambdaMxy", vbTextCompare) = 0 Or _
        StrComp(loadPath, "LambdaNMxy", vbTextCompare) = 0)
End Function

Private Function CapacityLoadPathScalesMy(ByVal loadPath As String) As Boolean
    CapacityLoadPathScalesMy = (StrComp(loadPath, "LambdaMy", vbTextCompare) = 0 Or _
        StrComp(loadPath, "LambdaMxy", vbTextCompare) = 0 Or _
        StrComp(loadPath, "LambdaNMxy", vbTextCompare) = 0)
End Function
