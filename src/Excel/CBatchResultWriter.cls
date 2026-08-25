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
' отчет по всем сочетаниям. Таблица разделена на блоки: запас по деформациям,
' запас по предельному моменту и расчет раскрытия трещин. Расчетные алгоритмы
' здесь не запускаются: класс только берет готовые данные CBatchSectionCalculator
' и записывает их на лист одним массивом, а коэффициенты запаса и a_crc задает
' формулами Excel, чтобы пользователь видел связь между исходными колонками.

Private Const TARGET_ROWS As Long = 31 ' Высота фиксированного диапазона batch summary.
Private Const TARGET_COLUMNS As Long = 46 ' Ширина новой сводки до общего минимального запаса.
Private Const META_ROWS As Long = 6 ' Верхние строки с общими сведениями о запуске.
Private Const GROUP_HEADER_ROW As Long = 7 ' Строка объединенных заголовков расчетных блоков.
Private Const SUBGROUP_HEADER_ROW As Long = 8 ' Строка внутренних подзаголовков расчетных блоков.
Private Const COLUMN_HEADER_ROW As Long = 9 ' Строка заголовков конкретных параметров.
Private Const FIRST_DATA_ROW As Long = 10 ' Первая строка сочетаний.

Private Const COL_ID As Long = 1
Private Const COL_COMMENT As Long = 2
Private Const COL_CALC_TYPE As Long = 3
Private Const COL_STATUS As Long = 4

Private Const COL_EPS_C_COMP As Long = 5
Private Const COL_EPS_C_TENS As Long = 6
Private Const COL_EPS_S_COMP As Long = 7
Private Const COL_EPS_S_TENS As Long = 8
Private Const COL_STATE_H As Long = 9
Private Const COL_STATE_C As Long = 10
Private Const COL_LIM_C_COMP As Long = 11
Private Const COL_LIM_C_TENS As Long = 12
Private Const COL_LIM_S_COMP As Long = 13
Private Const COL_LIM_S_TENS As Long = 14
Private Const COL_STRAIN_SF As Long = 15
Private Const COL_LIMIT_STATE As Long = 16

Private Const COL_MXY As Long = 17
Private Const COL_MXY_ULT As Long = 18
Private Const COL_MOMENT_H As Long = 19
Private Const COL_MOMENT_C As Long = 20
Private Const COL_MOMENT_SF As Long = 21

Private Const COL_CRACK_STATUS As Long = 22
Private Const COL_CRACK_FORMED As Long = 23
Private Const COL_CRACK_H As Long = 24
Private Const COL_CRACK_A As Long = 25
Private Const COL_CRACK_TENSION_DEPTH As Long = 26
Private Const COL_CRACK_EFFECTIVE_DEPTH As Long = 27
Private Const COL_CRACK_AS As Long = 28
Private Const COL_CRACK_ABT As Long = 29
Private Const COL_CRACK_DS As Long = 30
Private Const COL_CRACK_REBAR_COUNT As Long = 31
Private Const COL_CRACK_LS_RAW As Long = 32
Private Const COL_CRACK_LS As Long = 33
Private Const COL_CRACK_PHI1 As Long = 34
Private Const COL_CRACK_PHI2 As Long = 35
Private Const COL_CRACK_PHI3 As Long = 36
Private Const COL_CRACK_NCRC As Long = 37
Private Const COL_CRACK_LAMBDA_CRC As Long = 38
Private Const COL_CRACK_SIGMA_S As Long = 39
Private Const COL_CRACK_SIGMA_S_CRC As Long = 40
Private Const COL_CRACK_PSI As Long = 41
Private Const COL_CRACK_ES As Long = 42
Private Const COL_CRACK_WIDTH As Long = 43
Private Const COL_CRACK_ALLOWABLE As Long = 44
Private Const COL_CRACK_SF As Long = 45

Private Const COL_OVERALL_SF As Long = 46

Private Const HEADER_COLOR As Long = &HD9D9D9 ' Светло-серый фон заголовков.
Private Const SAFETY_COLOR As Long = &HE7E6E6 ' Серый фон столбцов коэффициентов запаса.
Private Const OVERALL_SAFETY_COLOR As Long = &HBFBFBF ' Более темный фон итогового минимального запаса.
Private Const NOT_CALCULATED_SHORT As String = "NotCalcul." ' Короткий статус для компактной таблицы Results.

' Записывает rngBatchSummary одним массивом и оформляет блоки.
' section нужен для геометрических величин h/x: в блоке деформаций они
' относятся к фактическому LC, а в блоке предельного момента - только к
' найденному capacity-состоянию. Если section не передан, эти поля пустые.
Public Sub WriteSummary(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, _
        Optional ByVal units As CUnitSystem = Nothing, Optional ByVal section As CSectionModel = Nothing)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3970, "CBatchResultWriter", "Книга Excel не передана."
    If batch Is Nothing Then Err.Raise vbObjectError + 3971, "CBatchResultWriter", "Пакетный расчетчик не передан."

    Dim target As Object
    Set target = SummaryTargetRange(workbook)
    PrepareSummaryTarget target

    Dim data() As Variant
    ReDim data(1 To TARGET_ROWS, 1 To TARGET_COLUMNS)

    PutMetaRows data, batch, units
    PutGroupHeaders data
    PutSubgroupHeaders data
    PutColumnHeaders data, units

    Dim i As Long
    For i = 1 To batch.Count
        PutCombination data, FIRST_DATA_ROW + i - 1, batch, i, units, section
    Next i

    WriteSummaryBody target, data
    FormatSummary workbook, target, batch.Count
End Sub

' Очищает старую сводку перед новым расчетом или ручной очисткой Results.
Public Sub ClearSummary(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3972, "CBatchResultWriter", "Книга Excel не передана."
    ClearSummaryBody SummaryTargetRange(workbook)
End Sub

' Обновляет итоговое время уже после завершения записи Results и схемы.
' Сама сводка формируется раньше, потому что Plot.LoadCase = Worst читает из
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

' Снимает объединения и очищает старую сводку до блочной записи массива.
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
    data(6, 1) = "Точка приложения нагрузки (относительно центра тяжести сечения)"
    data(6, 5) = "X=" & CStr(OutputLengthForSummary(batch.LoadReferenceOffsetX, units)) & " " & OutputLengthUnit(units)
    data(6, 8) = "Y=" & CStr(OutputLengthForSummary(batch.LoadReferenceOffsetY, units)) & " " & OutputLengthUnit(units)
End Sub

Private Sub PutGroupHeaders(ByRef data() As Variant)
    data(GROUP_HEADER_ROW, COL_EPS_C_COMP) = "Расчет несущей способности по относительным деформациям"
    data(GROUP_HEADER_ROW, COL_LIMIT_STATE) = "Расчет несущей способности по предельному моменту"
    data(GROUP_HEADER_ROW, COL_CRACK_STATUS) = "Расчет ширины раскрытия нормальных трещин"
    data(GROUP_HEADER_ROW, COL_OVERALL_SF) = "Минимальный запас"
End Sub

Private Sub PutSubgroupHeaders(ByRef data() As Variant)
    data(SUBGROUP_HEADER_ROW, COL_ID) = "Сочетание"
    data(SUBGROUP_HEADER_ROW, COL_EPS_C_COMP) = "Фактическое НДС от LC"
    data(SUBGROUP_HEADER_ROW, COL_LIM_C_COMP) = "Предельные деформации"
    data(SUBGROUP_HEADER_ROW, COL_STRAIN_SF) = "Запас"
    data(SUBGROUP_HEADER_ROW, COL_LIMIT_STATE) = "Предельное состояние"

    data(SUBGROUP_HEADER_ROW, COL_MXY) = "Моменты"
    data(SUBGROUP_HEADER_ROW, COL_MOMENT_H) = "Геометрия при MxyUltimate"
    data(SUBGROUP_HEADER_ROW, COL_MOMENT_SF) = "Запас"

    data(SUBGROUP_HEADER_ROW, COL_CRACK_STATUS) = "Фактическое НДС от LC"
    data(SUBGROUP_HEADER_ROW, COL_CRACK_AS) = "Арматура и бетон"
    data(SUBGROUP_HEADER_ROW, COL_CRACK_LS_RAW) = "Расстояние между трещинами"
    data(SUBGROUP_HEADER_ROW, COL_CRACK_PHI1) = "Коэффициенты phi"
    data(SUBGROUP_HEADER_ROW, COL_CRACK_NCRC) = "Расчет psi_s"
    data(SUBGROUP_HEADER_ROW, COL_CRACK_ES) = "Модуль арматуры"
    data(SUBGROUP_HEADER_ROW, COL_CRACK_WIDTH) = "Проверка a_crc"
    data(SUBGROUP_HEADER_ROW, COL_OVERALL_SF) = "Итог"
End Sub

Private Sub PutColumnHeaders(ByRef data() As Variant, ByVal units As CUnitSystem)
    data(COLUMN_HEADER_ROW, COL_ID) = "CombinationID (сочетание)"
    data(COLUMN_HEADER_ROW, COL_COMMENT) = "Comment (комментарий)"
    data(COLUMN_HEADER_ROW, COL_CALC_TYPE) = "CalculationType (группа)"
    data(COLUMN_HEADER_ROW, COL_STATUS) = "Status (статус)"

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
    data(COLUMN_HEADER_ROW, COL_STRAIN_SF) = "StrainSafetyFactor (запас по деформациям)"
    data(COLUMN_HEADER_ROW, COL_LIMIT_STATE) = "CapacityLimitState (предельное состояние)"

    data(COLUMN_HEADER_ROW, COL_MXY) = "Mxy, " & OutputMomentUnit(units) & " (действующий момент)"
    data(COLUMN_HEADER_ROW, COL_MXY_ULT) = "MxyUltimate, " & OutputMomentUnit(units) & " (несущий момент)"
    data(COLUMN_HEADER_ROW, COL_MOMENT_H) = "h, " & OutputLengthUnit(units) & " (высота при MxyUltimate)"
    data(COLUMN_HEADER_ROW, COL_MOMENT_C) = "x, " & OutputLengthUnit(units) & " (сжатая зона при MxyUltimate)"
    data(COLUMN_HEADER_ROW, COL_MOMENT_SF) = "MomentSafetyFactor (запас по моменту)"

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
    data(COLUMN_HEADER_ROW, COL_CRACK_PHI2) = "phi2 (профиль арматуры)"
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

    data(COLUMN_HEADER_ROW, COL_OVERALL_SF) = "MinSafetyFactor (общий запас)"
End Sub

Private Sub PutCombination(ByRef data() As Variant, ByVal rowIndex As Long, ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal units As CUnitSystem, ByVal section As CSectionModel)
    If rowIndex > UBound(data, 1) Then Exit Sub

    data(rowIndex, COL_ID) = batch.CombinationID(index)
    data(rowIndex, COL_COMMENT) = batch.CombinationName(index)
    data(rowIndex, COL_CALC_TYPE) = batch.CalculationType(index)
    data(rowIndex, COL_STATUS) = ShortStatus(batch.Status(index))

    data(rowIndex, COL_EPS_C_COMP) = batch.MinConcreteStrain(index)
    data(rowIndex, COL_EPS_C_TENS) = batch.MaxConcreteStrain(index)
    data(rowIndex, COL_EPS_S_COMP) = batch.MinSteelStrain(index)
    data(rowIndex, COL_EPS_S_TENS) = batch.MaxSteelStrain(index)
    PutCurrentStateGeometry data, rowIndex, batch, index, units, section
    data(rowIndex, COL_LIM_C_COMP) = batch.ConcreteCompressionLimitForCombination(index)
    If batch.ConcreteTensionLimitEnabledForCombination(index) Then data(rowIndex, COL_LIM_C_TENS) = batch.ConcreteTensionLimitForCombination(index)
    data(rowIndex, COL_LIM_S_COMP) = batch.SteelCompressionLimitForCombination(index)
    data(rowIndex, COL_LIM_S_TENS) = batch.SteelTensionLimitForCombination(index)
    data(rowIndex, COL_STRAIN_SF) = StrainSafetyFormula(rowIndex, batch, index)
    data(rowIndex, COL_LIMIT_STATE) = ShortStatus(batch.CapacityLimitState(index))

    data(rowIndex, COL_MXY) = OutputMomentMagnitude(Sqr(batch.UserMx(index) ^ 2 + batch.UserMy(index) ^ 2), units)
    If batch.MomentUltimate(index) > 0# Then
        data(rowIndex, COL_MXY_ULT) = OutputMomentMagnitude(batch.MomentUltimate(index), units)
        data(rowIndex, COL_MOMENT_SF) = MomentSafetyFormula(rowIndex)
    End If
    PutCapacityMomentGeometry data, rowIndex, batch, index, units, section

    data(rowIndex, COL_CRACK_STATUS) = ShortStatus(batch.CrackStatus(index))
    data(rowIndex, COL_CRACK_NCRC) = NOT_CALCULATED_SHORT
    If StrComp(batch.CrackStatus(index), "OK", vbTextCompare) = 0 Then
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

    data(rowIndex, COL_OVERALL_SF) = OverallSafetyFormula(rowIndex)
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
Private Sub PutCapacityMomentGeometry(ByRef data() As Variant, ByVal rowIndex As Long, ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal units As CUnitSystem, ByVal section As CSectionModel)
    If section Is Nothing Then Exit Sub
    If batch.MomentUltimate(index) <= 0# Then Exit Sub
    If Not batch.CapacityStateAvailable(index) Then Exit Sub

    Dim depthH As Double
    Dim compressedDepth As Double
    CalculateDepthsPerpendicularToNeutral section, batch.CapacityEpsilon0(index), batch.CapacityKappaX(index), batch.CapacityKappaY(index), _
        depthH, compressedDepth
    If depthH > 0# Then data(rowIndex, COL_MOMENT_H) = OutputLength(depthH, units)
    If compressedDepth >= 0# Then data(rowIndex, COL_MOMENT_C) = OutputLength(compressedDepth, units)
End Sub

' Определяет габарит h и глубину сжатой зоны c вдоль нормали к нейтральной линии.
' Для прямоугольных бетонных волокон учитывается их width/height; для импортных
' элементов без размеров используется эквивалентный по площади квадрат.
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
    Dim i As Long
    For i = 1 To section.ConcreteCount
        Dim halfProjection As Double
        halfProjection = 0.5 * (Abs(nx) * ConcreteWidth(section, i) + Abs(ny) * ConcreteHeight(section, i))
        Dim centerS As Double
        centerS = nx * section.ConcreteX(i) + ny * section.ConcreteY(i)
        If i = 1 Then
            minS = centerS - halfProjection
            maxS = centerS + halfProjection
        Else
            minS = MinDouble(minS, centerS - halfProjection)
            maxS = MaxDouble(maxS, centerS + halfProjection)
        End If
    Next i

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

Private Function StrainSafetyFormula(ByVal rowIndex As Long, ByVal batch As CBatchSectionCalculator, ByVal index As Long) As String
    Dim actualColumn As Long
    Dim limitColumn As Long
    Dim bestSafety As Double

    PickStrainSafetyCandidate batch.MinConcreteStrain(index), batch.ConcreteCompressionLimitForCombination(index), _
        COL_EPS_C_COMP, COL_LIM_C_COMP, bestSafety, actualColumn, limitColumn

    If batch.ConcreteTensionLimitEnabledForCombination(index) Then
        PickStrainSafetyCandidate batch.MaxConcreteStrain(index), batch.ConcreteTensionLimitForCombination(index), _
            COL_EPS_C_TENS, COL_LIM_C_TENS, bestSafety, actualColumn, limitColumn
    End If

    PickStrainSafetyCandidate batch.MinSteelStrain(index), batch.SteelCompressionLimitForCombination(index), _
        COL_EPS_S_COMP, COL_LIM_S_COMP, bestSafety, actualColumn, limitColumn
    PickStrainSafetyCandidate batch.MaxSteelStrain(index), batch.SteelTensionLimitForCombination(index), _
        COL_EPS_S_TENS, COL_LIM_S_TENS, bestSafety, actualColumn, limitColumn

    If actualColumn > 0 And limitColumn > 0 Then
        StrainSafetyFormula = "=IFERROR(ABS(" & CellAddress(rowIndex, limitColumn) & "/" & _
            CellAddress(rowIndex, actualColumn) & "),"""")"
    End If
End Function

Private Function MomentSafetyFormula(ByVal rowIndex As Long) As String
    MomentSafetyFormula = "=IFERROR(" & CellAddress(rowIndex, COL_MXY_ULT) & "/" & _
        CellAddress(rowIndex, COL_MXY) & ","""")"
End Function

Private Sub PickStrainSafetyCandidate(ByVal actualStrain As Double, ByVal limitStrain As Double, _
        ByVal actualColumn As Long, ByVal limitColumn As Long, ByRef bestSafety As Double, _
        ByRef bestActualColumn As Long, ByRef bestLimitColumn As Long)
    If actualStrain = 0# Then Exit Sub
    If limitStrain = 0# Then Exit Sub
    If Sgn(actualStrain) <> Sgn(limitStrain) Then Exit Sub

    Dim safety As Double
    safety = Abs(limitStrain / actualStrain)
    If safety <= 0# Then Exit Sub

    If bestSafety <= 0# Or safety < bestSafety Then
        bestSafety = safety
        bestActualColumn = actualColumn
        bestLimitColumn = limitColumn
    End If
End Sub

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

Private Function OverallSafetyFormula(ByVal rowIndex As Long) As String
    OverallSafetyFormula = "=IFERROR(AGGREGATE(15,6,CHOOSE({1,2,3}," & _
        "IF(" & CellAddress(rowIndex, COL_STRAIN_SF) & ">0," & CellAddress(rowIndex, COL_STRAIN_SF) & ",NA())," & _
        "IF(" & CellAddress(rowIndex, COL_MOMENT_SF) & ">0," & CellAddress(rowIndex, COL_MOMENT_SF) & ",NA())," & _
        "IF(" & CellAddress(rowIndex, COL_CRACK_SF) & ">0," & CellAddress(rowIndex, COL_CRACK_SF) & ",NA())),1),"""")"
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

Private Sub FormatSummary(ByVal workbook As Object, ByVal target As Object, ByVal combinationCount As Long)
    target.UnMerge
    target.Interior.ColorIndex = -4142
    target.Borders.LineStyle = -4142

    target.HorizontalAlignment = -4131
    target.VerticalAlignment = -4108
    target.WrapText = False
    target.Font.Name = "Arial"
    target.Font.Size = 9

    target.Rows.Item(1).Font.Bold = True
    target.Rows.Item(GROUP_HEADER_ROW).Font.Bold = True
    target.Rows.Item(SUBGROUP_HEADER_ROW).Font.Bold = True
    target.Rows.Item(COLUMN_HEADER_ROW).Font.Bold = True
    target.Rows.Item(GROUP_HEADER_ROW).Interior.Color = HEADER_COLOR
    target.Rows.Item(SUBGROUP_HEADER_ROW).Interior.Color = HEADER_COLOR
    target.Rows.Item(COLUMN_HEADER_ROW).Interior.Color = HEADER_COLOR
    target.Rows.Item(1).Interior.Color = HEADER_COLOR

    MergeSummaryHeader target, GROUP_HEADER_ROW, COL_EPS_C_COMP, COL_STRAIN_SF
    MergeSummaryHeader target, GROUP_HEADER_ROW, COL_LIMIT_STATE, COL_MOMENT_SF
    MergeSummaryHeader target, GROUP_HEADER_ROW, COL_CRACK_STATUS, COL_CRACK_SF
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_ID, COL_STATUS
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_EPS_C_COMP, COL_STATE_C
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_LIM_C_COMP, COL_LIM_S_TENS
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_MXY, COL_MXY_ULT
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_MOMENT_H, COL_MOMENT_C
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_CRACK_STATUS, COL_CRACK_EFFECTIVE_DEPTH
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_CRACK_AS, COL_CRACK_REBAR_COUNT
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_CRACK_LS_RAW, COL_CRACK_LS
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_CRACK_PHI1, COL_CRACK_PHI3
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_CRACK_NCRC, COL_CRACK_PSI
    MergeSummaryHeader target, SUBGROUP_HEADER_ROW, COL_CRACK_WIDTH, COL_CRACK_SF

    target.Rows.Item(COLUMN_HEADER_ROW).HorizontalAlignment = -4131
    target.Rows.Item(GROUP_HEADER_ROW).HorizontalAlignment = -4108
    target.Rows.Item(SUBGROUP_HEADER_ROW).HorizontalAlignment = -4108
    target.Rows.Item(1).HorizontalAlignment = -4131
    target.Offset(0, 0).Resize(META_ROWS, TARGET_COLUMNS).HorizontalAlignment = -4131

    Dim dataRows As Long
    dataRows = TARGET_ROWS - FIRST_DATA_ROW + 1
    target.Offset(FIRST_DATA_ROW - 1, 0).Resize(dataRows, TARGET_COLUMNS).HorizontalAlignment = -4108
    target.Offset(FIRST_DATA_ROW - 1, COL_COMMENT - 1).Resize(dataRows, 1).HorizontalAlignment = -4152
    target.Offset(FIRST_DATA_ROW - 1, COL_STRAIN_SF - 1).Resize(dataRows, 1).Interior.Color = SAFETY_COLOR
    target.Offset(FIRST_DATA_ROW - 1, COL_MOMENT_SF - 1).Resize(dataRows, 1).Interior.Color = SAFETY_COLOR
    target.Offset(FIRST_DATA_ROW - 1, COL_CRACK_SF - 1).Resize(dataRows, 1).Interior.Color = SAFETY_COLOR
    target.Offset(FIRST_DATA_ROW - 1, COL_OVERALL_SF - 1).Resize(dataRows, 1).Interior.Color = OVERALL_SAFETY_COLOR

    Dim borderedArea As Object
    Set borderedArea = target.Offset(GROUP_HEADER_ROW - 1, 0).Resize(TARGET_ROWS - GROUP_HEADER_ROW + 1, TARGET_COLUMNS)
    borderedArea.Borders.LineStyle = 1
    borderedArea.Borders.Color = RGB(0, 0, 0)
    borderedArea.Borders.Weight = 2
    target.Offset(0, 0).Resize(1, TARGET_COLUMNS).NumberFormat = "@"
End Sub

Private Sub MergeSummaryHeader(ByVal target As Object, ByVal rowIndex As Long, ByVal firstColumn As Long, ByVal lastColumn As Long)
    Dim area As Object
    Set area = target.Offset(rowIndex - 1, firstColumn - 1).Resize(1, lastColumn - firstColumn + 1)
    area.Merge
End Sub

' Сокращает длинный служебный статус только в пользовательской сводке.
' В расчетных объектах и подробных Results остается исходный текст статуса.
Private Function ShortStatus(ByVal statusText As String) As String
    ShortStatus = Replace$(statusText, "NotCalculated", NOT_CALCULATED_SHORT, 1, -1, vbTextCompare)
End Function

Private Function ConcreteWidth(ByVal section As CSectionModel, ByVal index As Long) As Double
    ConcreteWidth = section.ConcreteWidth(index)
    If ConcreteWidth <= 0# Then ConcreteWidth = Sqr(section.ConcreteArea(index))
End Function

Private Function ConcreteHeight(ByVal section As CSectionModel, ByVal index As Long) As Double
    ConcreteHeight = section.ConcreteHeight(index)
    If ConcreteHeight <= 0# Then ConcreteHeight = Sqr(section.ConcreteArea(index))
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

Private Function OutputForceUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputForceUnit = "N" Else OutputForceUnit = units.OutputForceUnit
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
