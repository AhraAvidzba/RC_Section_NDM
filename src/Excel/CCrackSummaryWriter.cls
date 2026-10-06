VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCrackSummaryWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Детальная таблица расчета нормальных и продольных трещин на листе Results
' ==========================================================================
' Класс пишет отдельный блок от якоря rngCrackSummaryAnchor. Расчет здесь не
' выполняется: writer берет уже сохраненные в CBatchSectionCalculator нагрузки,
' стадии трещинообразования, параметры ширины раскрытия и проверку продольных
' трещин, переводит их в пользовательские единицы и оформляет таблицу.
'
' Верхний rngBatchSummary остается компактной сводкой, а этот блок раскрывает
' расчет трещин по сочетаниям. Пустой столбец между левой частью и подробным
' блоком специально остается без рамок и заливки, как в таблице устойчивости.

Private Const HEADER_ROWS As Long = 4 ' Сколько строк шапки строится над якорем первой строки LC.
Private Const TOTAL_COLUMNS As Long = 72 ' Относительная ширина Formation/SP35/SP63/продольных с разделителями.

Private Const COL_LEFT_FIRST As Long = 1
Private Const COL_LEFT_LAST As Long = 9
Private Const COL_SEPARATOR As Long = 10
Private Const COL_MAIN_FIRST As Long = 11
Private Const COL_MAIN_LAST As Long = TOTAL_COLUMNS

Private Const COL_ID As Long = 1
Private Const COL_RESULT_COMMENT As Long = 2
Private Const COL_STATUS As Long = 3
Private Const COL_REF_N As Long = 4
Private Const COL_REF_MX As Long = 5
Private Const COL_REF_MY As Long = 6
Private Const COL_LOAD_N As Long = 7
Private Const COL_LOAD_MX As Long = 8
Private Const COL_LOAD_MY As Long = 9

Private Const COL_ARED As Long = 11
Private Const COL_RBT_SER As Long = 12
Private Const COL_NCRC As Long = 13
Private Const COL_LAMBDA_NCRC As Long = 14
Private Const COL_FORMATION_METHOD As Long = 15
Private Const COL_FORMATION_NCRC As Long = 16
Private Const COL_MCRC As Long = 17
Private Const COL_LAMBDA_MCRC As Long = 18
Private Const COL_PRE_EBC As Long = COL_LAMBDA_MCRC + 1
Private Const COL_PRE_EBT As Long = COL_PRE_EBC + 1
Private Const COL_PRE_CRACK_STATUS As Long = COL_PRE_EBT + 1
Private Const COL_POST_EBC As Long = COL_PRE_CRACK_STATUS + 1
Private Const COL_POST_EBT As Long = COL_POST_EBC + 1
Private Const COL_POST_CRACK_STATUS As Long = COL_POST_EBT + 1
Private Const COL_FORMATION_STATE As Long = COL_POST_CRACK_STATUS + 1
Private Const COL_CURRENT_EBC As Long = COL_FORMATION_STATE + 1
Private Const COL_CURRENT_EBT As Long = COL_CURRENT_EBC + 1
Private Const COL_CRACK_STATE_DIAGEXT As Long = COL_CURRENT_EBT + 1
Private Const COL_CRACK_STATE_STATUS As Long = COL_CRACK_STATE_DIAGEXT + 1
Private Const COL_FORMATION_SEPARATOR As Long = COL_CRACK_STATE_STATUS + 1
Private Const COL_SP35_AR As Long = COL_FORMATION_SEPARATOR + 1
Private Const COL_SP35_BETA As Long = COL_SP35_AR + 1
Private Const COL_SP35_N As Long = COL_SP35_BETA + 1
Private Const COL_SP35_D As Long = COL_SP35_N + 1
Private Const COL_SP35_RR As Long = COL_SP35_D + 1
Private Const COL_SP35_PSI As Long = COL_SP35_RR + 1
Private Const COL_SP35_REBARS As Long = COL_SP35_PSI + 1
Private Const COL_SP35_NREBARS As Long = COL_SP35_REBARS + 1
Private Const COL_SP35_SIGMA As Long = COL_SP35_NREBARS + 1
Private Const COL_SP35_ES As Long = COL_SP35_SIGMA + 1
Private Const COL_SP35_ACRC As Long = COL_SP35_ES + 1
Private Const COL_SP35_ALLOWABLE As Long = COL_SP35_ACRC + 1
Private Const COL_SP35_RATIO As Long = COL_SP35_ALLOWABLE + 1
Private Const COL_SP35_STATUS As Long = COL_SP35_RATIO + 1
Private Const COL_SP35_SEPARATOR As Long = COL_SP35_STATUS + 1
Private Const COL_XT As Long = COL_SP35_SEPARATOR + 1
Private Const COL_H As Long = COL_XT + 1
Private Const COL_AS_COVER As Long = COL_H + 1
Private Const COL_HBT As Long = COL_AS_COVER + 1
Private Const COL_ABT As Long = COL_HBT + 1
Private Const COL_AS As Long = COL_ABT + 1
Private Const COL_DS_EQ As Long = COL_AS + 1
Private Const COL_LS0 As Long = COL_DS_EQ + 1
Private Const COL_LS As Long = COL_LS0 + 1
Private Const COL_PHI1 As Long = COL_LS + 1
Private Const COL_PHI2 As Long = COL_PHI1 + 1
Private Const COL_PHI3 As Long = COL_PHI2 + 1
Private Const COL_PSI_S As Long = COL_PHI3 + 1
Private Const COL_REBARS As Long = COL_PSI_S + 1
Private Const COL_NREBARS As Long = COL_REBARS + 1
Private Const COL_SIGMA_S_CRC As Long = COL_NREBARS + 1
Private Const COL_SIGMA_S As Long = COL_SIGMA_S_CRC + 1
Private Const COL_ES As Long = COL_SIGMA_S + 1
Private Const COL_ACRC As Long = COL_ES + 1
Private Const COL_ACRC_ULT As Long = COL_ACRC + 1
Private Const COL_ACRC_RATIO As Long = COL_ACRC_ULT + 1
Private Const COL_ACRC_STATUS As Long = COL_ACRC_RATIO + 1
Private Const COL_SP63_SEPARATOR As Long = COL_ACRC_STATUS + 1
Private Const COL_SIGMA_B_MAX As Long = COL_SP63_SEPARATOR + 1
Private Const COL_RB_MC2 As Long = COL_SIGMA_B_MAX + 1
Private Const COL_LONG_RATIO As Long = COL_RB_MC2 + 1
Private Const COL_LONG_STATUS As Long = COL_LONG_RATIO + 1

Private Const HEADER_COLOR As Long = &HD9D9D9 ' Серый фон шапки, единый для таблиц Results.
Private Const SMALL_VALUE As Double = 0.000000001 ' Порог только для пустых расчетных полей writer-а.
Private Const HEADER_NOTES_ROW As Long = 3 ' Строка коротких пояснений над нижними обозначениями.
Private Const HEADER_LABEL_ROW As Long = 4 ' Строка итоговых обозначений с единицами.

' Возвращает число строк шапки, расположенных выше rngCrackSummaryAnchor.
Public Function HeaderRowsAboveAnchor() As Long
    HeaderRowsAboveAnchor = HEADER_ROWS
End Function

' Возвращает полную высоту блока трещин с учетом шапки и строк сочетаний.
Public Function RequiredRows(ByVal combinationCount As Long) As Long
    If combinationCount < 1 Then combinationCount = 1
    RequiredRows = HEADER_ROWS + combinationCount
End Function

' Возвращает число строк данных, которое нужно оформить по текущему диапазону
' сочетаний. Пустые доступные строки тоже получают рамки, как в batch summary.
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

' Возвращает полную высоту блока для фактической книги.
Public Function RequiredRowsForWorkbook(ByVal workbook As Object, ByVal combinationCount As Long) As Long
    RequiredRowsForWorkbook = RequiredRows(RequiredDataRowsForWorkbook(workbook, combinationCount))
End Function

' Возвращает ширину блока вместе с пустым разделительным столбцом.
Public Function RequiredColumns() As Long
    RequiredColumns = TOTAL_COLUMNS
End Function

' Записывает детальную таблицу трещин от rngCrackSummaryAnchor.
Public Sub WriteSummary(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, _
        Optional ByVal units As CUnitSystem = Nothing)
    On Error GoTo Failed
    If workbook Is Nothing Then Err.Raise vbObjectError + 3988, "CCrackSummaryWriter", "Книга Excel не передана."
    If batch Is Nothing Then Err.Raise vbObjectError + 3989, "CCrackSummaryWriter", "Пакетный расчетчик не передан."

    Dim target As Object
    Set target = TargetRange(workbook, RequiredDataRowsForWorkbook(workbook, batch.Count))
    PrepareTarget target
    WriteBlockTitle target, "Расчет нормальных и продольных трещин"
    WriteHeaders target, units
    WriteData target, batch, units
    FormatTable target, batch, units
    Exit Sub

Failed:
    Err.Raise Err.Number, "CCrackSummaryWriter.WriteSummary", Err.Description
End Sub

' Полностью очищает старый блок трещин вместе с объединениями и форматами.
Public Sub ClearSummary(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3990, "CCrackSummaryWriter", "Книга Excel не передана."
    PrepareTarget ClearTargetRange(workbook)
End Sub

' Возвращает полный диапазон текущей таблицы: шапка над якорем и строки LC.
Private Function TargetRange(ByVal workbook As Object, Optional ByVal dataRows As Long = 1) As Object
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    If anchor.Row <= HEADER_ROWS Then Err.Raise vbObjectError + 3991, "CCrackSummaryWriter", _
        "rngCrackSummaryAnchor должен иметь место для шапки над собой."
    Set TargetRange = anchor.Offset(-HEADER_ROWS, 0).Resize(RequiredRows(dataRows), TOTAL_COLUMNS)
End Function

' Возвращает диапазон безопасной очистки до следующего нижнего блока Results.
Private Function ClearTargetRange(ByVal workbook As Object) As Object
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    If anchor.Row <= HEADER_ROWS Then Err.Raise vbObjectError + 3991, "CCrackSummaryWriter", _
        "rngCrackSummaryAnchor должен иметь место для шапки над собой."

    Dim topRow As Long
    topRow = anchor.Row - HEADER_ROWS

    Dim rowsToClear As Long
    rowsToClear = RequiredRows(1)
    On Error GoTo UseDefaultHeight
    Dim stabilityWriter As CStabilitySummaryWriter
    Set stabilityWriter = New CStabilitySummaryWriter
    Dim stabilityAnchor As Object
    Set stabilityAnchor = workbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange
    Dim stabilityTopRow As Long
    stabilityTopRow = stabilityAnchor.Row - stabilityWriter.HeaderRowsAboveAnchor
    If stabilityTopRow > topRow Then rowsToClear = stabilityTopRow - topRow

UseDefaultHeight:
    If topRow + rowsToClear - 1 > anchor.Worksheet.Rows.Count Then rowsToClear = anchor.Worksheet.Rows.Count - topRow + 1
    Set ClearTargetRange = anchor.Worksheet.Cells(topRow, anchor.Column).Resize(rowsToClear, TOTAL_COLUMNS)
End Function

' Снимает старые объединения и очищает значения/форматы перед новой записью.
Private Sub PrepareTarget(ByVal target As Object)
    ClearBlockTitleRow target
    target.UnMerge
    target.Clear
End Sub

' Записывает отдельную строку-название над detailed-блоком трещин.
Private Sub WriteBlockTitle(ByVal target As Object, ByVal titleText As String)
    If target.Row <= 1 Then Exit Sub

    Dim titleRow As Long
    titleRow = target.Row - 1
    target.Worksheet.Cells(titleRow, target.Column).Resize(1, TOTAL_COLUMNS).Interior.Color = RGB(189, 215, 238)
    target.Worksheet.Rows.Item(titleRow).RowHeight = 15
    With target.Worksheet.Cells(titleRow, target.Column)
        .Value2 = titleText
        .Font.Name = "Arial"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(255, 0, 0)
        .HorizontalAlignment = -4131
        .VerticalAlignment = -4108
        .WrapText = False
    End With
End Sub

' Очищает название и заливку только в ширине блока трещин. Ячейки справа
' не относятся к этой таблице и сохраняются при повторной записи результатов.
Private Sub ClearBlockTitleRow(ByVal target As Object)
    If target.Row <= 1 Then Exit Sub
    target.Worksheet.Cells(target.Row - 1, target.Column).Resize(1, TOTAL_COLUMNS).Clear
    target.Worksheet.Rows.Item(target.Row - 1).RowHeight = 15
End Sub

' Записывает многоуровневую шапку таблицы.
Private Sub WriteHeaders(ByVal target As Object, ByVal units As CUnitSystem)
    Dim headers() As Variant
    ReDim headers(1 To HEADER_ROWS, 1 To TOTAL_COLUMNS)

    headers(1, COL_ID) = "Combination ID"
    headers(1, COL_RESULT_COMMENT) = "ResultComment"
    headers(1, COL_STATUS) = "общий статус"
    headers(1, COL_REF_N) = "относительно центральных осей приведенного сечения"
    headers(1, COL_LOAD_N) = "в точке приложения нагрузок"

    headers(1, COL_ARED) = "Начало трещинообразования"
    headers(2, COL_ARED) = "Осевое растяжение"
    headers(2, COL_FORMATION_METHOD) = "Изгиб N + M"
    headers(2, COL_PRE_EBC) = "PreCrackState (до трещины)"
    headers(2, COL_POST_EBC) = "PostCrackState (после трещины)"
    headers(1, COL_CURRENT_EBC) = "равновесное состояние при заданных нагрузках"
    headers(1, COL_SP35_AR) = "нормальные трещины по СП 35"
    headers(2, COL_SP35_AR) = "Коэффициент раскрытия трещин"
    headers(2, COL_SP35_REBARS) = "Напряжения"
    headers(2, COL_SP35_ACRC) = "ширина раскрытия нормальных трещин"
    headers(1, COL_XT) = "нормальные трещины по СП 63"
    headers(2, COL_XT) = "данные для расстояния между трещинами ls"
    headers(2, COL_DS_EQ) = "Расстояние между трещинами ls"
    headers(2, COL_PHI1) = "Коэффициенты"
    headers(2, COL_REBARS) = "Напряжения"
    headers(2, COL_ACRC) = "ширина раскрытия нормальных трещин"
    headers(1, COL_SIGMA_B_MAX) = "продольные трещины"

    PutHeaderNotes headers, units
    PutColumnLabels headers, units
    target.Resize(HEADER_ROWS, TOTAL_COLUMNS).Value2 = headers
    MergeHeaders target
End Sub

' Заполняет строку пояснений над нижним ярусом заголовков.
Private Sub PutHeaderNotes(ByRef headers() As Variant, ByVal units As CUnitSystem)
    headers(HEADER_NOTES_ROW, COL_PRE_EBC) = "фактическая наибольшая деформация сжатия бетона до трещины"
    headers(HEADER_NOTES_ROW, COL_PRE_EBT) = "фактическая наибольшая деформация растяжения бетона до трещины"
    headers(HEADER_NOTES_ROW, COL_POST_EBC) = "фактическая деформация сжатия после трещины"
    headers(HEADER_NOTES_ROW, COL_POST_EBT) = "растяжение после трещины; пусто при выключенном растянутом бетоне"
    headers(HEADER_NOTES_ROW, COL_CURRENT_EBC) = "фактическая деформация сжатия при нагрузках сочетания"
    headers(HEADER_NOTES_ROW, COL_CURRENT_EBT) = "растяжение при текущих нагрузках; пусто при выключенном растянутом бетоне"
    headers(HEADER_NOTES_ROW, COL_SP35_AR) = "площадь в " & OutputAreaUnit(units) & "; критическая область с вычетом всех проемов"
    headers(HEADER_NOTES_ROW, COL_SP35_BETA) = "коэффициенты групп каждого расчетного ряда"
    headers(HEADER_NOTES_ROW, COL_SP35_N) = "число реальных стержней каждой группы ряда"
    headers(HEADER_NOTES_ROW, COL_SP35_D) = "диаметры в " & OutputLengthUnit(units) & "; все стержни по рядам"
    headers(HEADER_NOTES_ROW, COL_SP35_RR) = "радиус армирования в " & OutputLengthUnit(units) & ": Ar / сумма beta*nd"
    headers(HEADER_NOTES_ROW, COL_SP35_PSI) = "коэффициент раскрытия в " & OutputLengthUnit(units) & "; переведен из см"
    headers(HEADER_NOTES_ROW, COL_SP35_REBARS) = "все включенные стержни каждого ряда критической области"
    headers(HEADER_NOTES_ROW, COL_SP35_NREBARS) = "общее число стержней критической области"
    headers(HEADER_NOTES_ROW, COL_SP35_SIGMA) = "напряжение в " & OutputStressUnit(units) & "; критическая группа"
    headers(HEADER_NOTES_ROW, COL_SP35_ES) = "модуль арматуры в " & OutputStressUnit(units) & "; текущая модель"
    headers(HEADER_NOTES_ROW, COL_SP35_ACRC) = "раскрытие в " & OutputLengthUnit(units) & "; максимум рассчитанных якорей"
    headers(HEADER_NOTES_ROW, COL_SP35_ALLOWABLE) = "допускаемая ширина в " & OutputLengthUnit(units) & "; настройка пользователя"
    headers(HEADER_NOTES_ROW, COL_SP35_RATIO) = "запас: допускаемая ширина / рассчитанная ширина"
    headers(HEADER_NOTES_ROW, COL_SP35_STATUS) = "статус собственной проверки ширины по общей policy"
    headers(HEADER_NOTES_ROW, COL_RESULT_COMMENT) = "пояснение расчетного статуса блока трещин"
    headers(HEADER_NOTES_ROW, COL_REF_N) = "нагрузки, переданные в расчет трещин, в центре приведенного сечения"
    headers(HEADER_NOTES_ROW, COL_REF_MX) = "момент Mx там же, вокруг глобальной оси X/Y принятой системы"
    headers(HEADER_NOTES_ROW, COL_REF_MY) = "момент My там же, вокруг глобальной оси X/Y принятой системы"
    headers(HEADER_NOTES_ROW, COL_LOAD_N) = "исходная продольная сила в пользовательской точке приложения"
    headers(HEADER_NOTES_ROW, COL_LOAD_MX) = "исходный Mx в пользовательской точке приложения"
    headers(HEADER_NOTES_ROW, COL_LOAD_MY) = "исходный My в пользовательской точке приложения"
    headers(HEADER_NOTES_ROW, COL_ARED) = "площадь приведенного сечения для Ncrc"
    headers(HEADER_NOTES_ROW, COL_RBT_SER) = "растяжение бетона II ГПС"
    headers(HEADER_NOTES_ROW, COL_FORMATION_METHOD) = "фактически использованный путь поиска точки образования трещины"
    headers(HEADER_NOTES_ROW, COL_NCRC) = "сила образования трещины при центральном растяжении"
    headers(HEADER_NOTES_ROW, COL_LAMBDA_NCRC) = "доля заданной растягивающей силы для Ncrc"
    headers(HEADER_NOTES_ROW, COL_FORMATION_NCRC) = "продольная сила в найденной точке образования трещины для выбранного пути"
    headers(HEADER_NOTES_ROW, COL_MCRC) = "модуль моментного вектора в найденной точке образования трещины"
    headers(HEADER_NOTES_ROW, COL_LAMBDA_MCRC) = "доля выбранного пути нагрузки до образования трещины"
    headers(HEADER_NOTES_ROW, COL_PRE_CRACK_STATUS) = "статус НДС непосредственно перед образованием нормальной трещины"
    headers(HEADER_NOTES_ROW, COL_POST_CRACK_STATUS) = "статус НДС сразу после образования нормальной трещины"
    headers(HEADER_NOTES_ROW, COL_FORMATION_STATE) = "итог: нормальная трещина образуется при заданном сочетании или нет"
    headers(HEADER_NOTES_ROW, COL_CRACK_STATE_DIAGEXT) = "yes - равновесие найдено через продолжение диаграммы; no - продолжение было разрешено, но не понадобилось; - - продолжение выключено в настройках"
    headers(HEADER_NOTES_ROW, COL_CRACK_STATE_STATUS) = "OK - равновесие при заданных нагрузках сошлось без продолжения диаграммы и все материалы остались в физическом диапазоне; FAIL/NumFail останавливают расчет трещин"
    headers(HEADER_NOTES_ROW, COL_XT) = "фактическая высота растянутой зоны"
    headers(HEADER_NOTES_ROW, COL_H) = "полная высота сечения по нормали к нейтральной линии"
    headers(HEADER_NOTES_ROW, COL_AS_COVER) = "расстояние от оси выбранного стержня до бетонного контура"
    headers(HEADER_NOTES_ROW, COL_HBT) = "принятая глубина зоны растянутого бетона"
    headers(HEADER_NOTES_ROW, COL_ABT) = "площадь растянутого бетона расчетной зоны"
    headers(HEADER_NOTES_ROW, COL_AS) = "площадь выбранной растянутой арматуры"
    headers(HEADER_NOTES_ROW, COL_DS_EQ) = "эквивалентный диаметр выбранных стержней"
    headers(HEADER_NOTES_ROW, COL_LS0) = "расстояние до нормативных ограничений"
    headers(HEADER_NOTES_ROW, COL_LS) = "принятое расстояние между трещинами"
    headers(HEADER_NOTES_ROW, COL_PHI1) = "коэффициент длительности действия нагрузки"
    headers(HEADER_NOTES_ROW, COL_PHI2) = "коэффициент профиля арматуры из Config"
    headers(HEADER_NOTES_ROW, COL_PHI3) = "коэффициент характера нагружения"
    headers(HEADER_NOTES_ROW, COL_PSI_S) = "коэффициент неравномерности деформаций арматуры"
    headers(HEADER_NOTES_ROW, COL_REBARS) = "стержни, использованные для As и " & ChrW$(&H3C3) & "s"
    headers(HEADER_NOTES_ROW, COL_NREBARS) = "количество выбранных стержней"
    headers(HEADER_NOTES_ROW, COL_SIGMA_S_CRC) = "напряжение выбранной арматуры сразу после образования трещины"
    headers(HEADER_NOTES_ROW, COL_SIGMA_S) = "напряжение выбранной арматуры при текущем НДС"
    headers(HEADER_NOTES_ROW, COL_ES) = "модуль упругости арматуры для формулы раскрытия трещины"
    headers(HEADER_NOTES_ROW, COL_ACRC) = "расчетная ширина раскрытия"
    headers(HEADER_NOTES_ROW, COL_ACRC_ULT) = "допустимая ширина раскрытия"
    headers(HEADER_NOTES_ROW, COL_ACRC_RATIO) = "запас: допустимая ширина делится на расчетную"
    headers(HEADER_NOTES_ROW, COL_ACRC_STATUS) = "статус проверки нормальных трещин"
    headers(HEADER_NOTES_ROW, COL_SIGMA_B_MAX) = "максимальное сжатие бетона для продольных трещин"
    headers(HEADER_NOTES_ROW, COL_RB_MC2) = "допустимое Rb,mc2"
    headers(HEADER_NOTES_ROW, COL_LONG_RATIO) = "запас: Rb,mc2 делится на сжатие бетона"
    headers(HEADER_NOTES_ROW, COL_LONG_STATUS) = "статус проверки продольных трещин"
End Sub

' Заполняет нижний ярус заголовков с единицами пользователя.
Private Sub PutColumnLabels(ByRef headers() As Variant, ByVal units As CUnitSystem)
    headers(HEADER_LABEL_ROW, COL_PRE_EBC) = ChrW$(&H3B5) & "bc"
    headers(HEADER_LABEL_ROW, COL_PRE_EBT) = ChrW$(&H3B5) & "bt"
    headers(HEADER_LABEL_ROW, COL_POST_EBC) = ChrW$(&H3B5) & "bc"
    headers(HEADER_LABEL_ROW, COL_POST_EBT) = ChrW$(&H3B5) & "bt"
    headers(HEADER_LABEL_ROW, COL_CURRENT_EBC) = ChrW$(&H3B5) & "bc"
    headers(HEADER_LABEL_ROW, COL_CURRENT_EBT) = ChrW$(&H3B5) & "bt"
    headers(HEADER_LABEL_ROW, COL_SP35_AR) = "Ar, " & OutputAreaUnit(units)
    headers(HEADER_LABEL_ROW, COL_SP35_BETA) = ChrW$(&H3B2)
    headers(HEADER_LABEL_ROW, COL_SP35_N) = "n"
    headers(HEADER_LABEL_ROW, COL_SP35_D) = "d, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_SP35_RR) = "Rr, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_SP35_PSI) = ChrW$(&H3C8) & ", " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_SP35_REBARS) = "rebars"
    headers(HEADER_LABEL_ROW, COL_SP35_NREBARS) = "nrebars"
    headers(HEADER_LABEL_ROW, COL_SP35_SIGMA) = ChrW$(&H3C3) & "s, " & OutputStressUnit(units)
    headers(HEADER_LABEL_ROW, COL_SP35_ES) = "Es, " & OutputStressUnit(units)
    headers(HEADER_LABEL_ROW, COL_SP35_ACRC) = "acrc, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_SP35_ALLOWABLE) = "acrc,ult, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_SP35_RATIO) = "acrc,ult/acrc"
    headers(HEADER_LABEL_ROW, COL_SP35_STATUS) = "статус"
    headers(HEADER_LABEL_ROW, COL_REF_N) = "N, " & OutputForceUnit(units)
    headers(HEADER_LABEL_ROW, COL_REF_MX) = "Mx, " & OutputMomentUnit(units)
    headers(HEADER_LABEL_ROW, COL_REF_MY) = "My, " & OutputMomentUnit(units)
    headers(HEADER_LABEL_ROW, COL_LOAD_N) = "N, " & OutputForceUnit(units)
    headers(HEADER_LABEL_ROW, COL_LOAD_MX) = "Mx, " & OutputMomentUnit(units)
    headers(HEADER_LABEL_ROW, COL_LOAD_MY) = "My, " & OutputMomentUnit(units)
    headers(HEADER_LABEL_ROW, COL_ARED) = "Ared, " & OutputAreaUnit(units)
    headers(HEADER_LABEL_ROW, COL_RBT_SER) = "Rbt,ser, " & OutputStressUnit(units)
    headers(HEADER_LABEL_ROW, COL_FORMATION_METHOD) = "Method"
    headers(HEADER_LABEL_ROW, COL_NCRC) = "Ncrc, " & OutputForceUnit(units)
    headers(HEADER_LABEL_ROW, COL_LAMBDA_NCRC) = ChrW$(&H3BB) & "crc"
    headers(HEADER_LABEL_ROW, COL_FORMATION_NCRC) = "Ncrc, " & OutputForceUnit(units)
    headers(HEADER_LABEL_ROW, COL_MCRC) = "Mxy,crc, " & OutputMomentUnit(units)
    headers(HEADER_LABEL_ROW, COL_LAMBDA_MCRC) = ChrW$(&H3BB) & "crc"
    headers(HEADER_LABEL_ROW, COL_PRE_CRACK_STATUS) = "равновесие"
    headers(HEADER_LABEL_ROW, COL_POST_CRACK_STATUS) = "равновесие"
    headers(HEADER_LABEL_ROW, COL_FORMATION_STATE) = "state"
    headers(HEADER_LABEL_ROW, COL_CRACK_STATE_DIAGEXT) = "ExtUsed"
    headers(HEADER_LABEL_ROW, COL_CRACK_STATE_STATUS) = "статус"
    headers(HEADER_LABEL_ROW, COL_XT) = "xt, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_H) = "h, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_AS_COVER) = "as, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_HBT) = "hbt, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_ABT) = "Abt, " & OutputAreaUnit(units)
    headers(HEADER_LABEL_ROW, COL_AS) = "As, " & OutputAreaUnit(units)
    headers(HEADER_LABEL_ROW, COL_DS_EQ) = "ds,eq, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_LS0) = "ls0, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_LS) = "ls, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_PHI1) = ChrW$(&H3C6) & "1"
    headers(HEADER_LABEL_ROW, COL_PHI2) = ChrW$(&H3C6) & "2"
    headers(HEADER_LABEL_ROW, COL_PHI3) = ChrW$(&H3C6) & "3"
    headers(HEADER_LABEL_ROW, COL_PSI_S) = ChrW$(&H3C8) & "s"
    headers(HEADER_LABEL_ROW, COL_REBARS) = "rebars"
    headers(HEADER_LABEL_ROW, COL_NREBARS) = "nrebars"
    headers(HEADER_LABEL_ROW, COL_SIGMA_S_CRC) = ChrW$(&H3C3) & "s,crc"
    headers(HEADER_LABEL_ROW, COL_SIGMA_S) = ChrW$(&H3C3) & "s"
    headers(HEADER_LABEL_ROW, COL_ES) = "Es, " & OutputStressUnit(units)
    headers(HEADER_LABEL_ROW, COL_ACRC) = "acrc, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_ACRC_ULT) = "acrc,ult, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_ACRC_RATIO) = "acrc,ult/acrc"
    headers(HEADER_LABEL_ROW, COL_ACRC_STATUS) = "статус"
    headers(HEADER_LABEL_ROW, COL_SIGMA_B_MAX) = ChrW$(&H3C3) & "b,max"
    headers(HEADER_LABEL_ROW, COL_RB_MC2) = "Rb,mc2"
    headers(HEADER_LABEL_ROW, COL_LONG_RATIO) = "Rb,mc2/" & ChrW$(&H3C3) & "b"
    headers(HEADER_LABEL_ROW, COL_LONG_STATUS) = "статус"
End Sub

' Объединяет ячейки многоуровневой шапки.
Private Sub MergeHeaders(ByVal target As Object)
    MergeRange target, 1, COL_ID, HEADER_LABEL_ROW, COL_ID
    MergeRange target, 1, COL_RESULT_COMMENT, HEADER_LABEL_ROW, COL_RESULT_COMMENT
    MergeRange target, 1, COL_STATUS, HEADER_LABEL_ROW, COL_STATUS
    MergeRange target, 1, COL_REF_N, 2, COL_REF_MY
    MergeRange target, 1, COL_LOAD_N, 2, COL_LOAD_MY
    MergeRange target, 1, COL_ARED, 1, COL_FORMATION_STATE
    MergeRange target, 2, COL_ARED, 2, COL_LAMBDA_NCRC
    MergeRange target, 2, COL_FORMATION_METHOD, 2, COL_LAMBDA_MCRC
    MergeRange target, 2, COL_PRE_EBC, 2, COL_PRE_CRACK_STATUS
    MergeRange target, 2, COL_POST_EBC, 2, COL_POST_CRACK_STATUS
    MergeRange target, 1, COL_CURRENT_EBC, 2, COL_CRACK_STATE_STATUS
    MergeRange target, 1, COL_SP35_AR, 1, COL_SP35_STATUS
    MergeRange target, 2, COL_SP35_AR, 2, COL_SP35_PSI
    MergeRange target, 2, COL_SP35_REBARS, 2, COL_SP35_ES
    MergeRange target, 2, COL_SP35_ACRC, 2, COL_SP35_STATUS
    MergeRange target, 1, COL_XT, 1, COL_ACRC_STATUS
    MergeRange target, 2, COL_XT, 2, COL_AS
    MergeRange target, 2, COL_DS_EQ, 2, COL_LS
    MergeRange target, 2, COL_PHI1, 2, COL_PSI_S
    MergeRange target, 2, COL_REBARS, 2, COL_SIGMA_S
    MergeRange target, 2, COL_ES, 2, COL_ACRC_STATUS
    MergeRange target, 1, COL_SIGMA_B_MAX, 2, COL_LONG_STATUS
End Sub

' Записывает строки сочетаний одним массивом.
Private Sub WriteData(ByVal target As Object, ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    Dim dataRows As Long
    dataRows = target.Rows.Count - HEADER_ROWS
    If dataRows <= 0 Then Exit Sub

    Dim data() As Variant
    ReDim data(1 To dataRows, 1 To TOTAL_COLUMNS)

    Dim i As Long
    For i = 1 To batch.Count
        PutCombination data, OutputDataOffset(batch, i, dataRows), batch, i, units
    Next i
    target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).Value2 = data
End Sub

' Заполняет одну строку сочетания фактически рассчитанными величинами трещин.
Private Sub PutCombination(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal units As CUnitSystem)
    Dim policy As CResultStatusPolicy
    Set policy = StatusPolicy()

    data(rowIndex, COL_ID) = batch.CombinationID(index)
    data(rowIndex, COL_RESULT_COMMENT) = MetaComment(batch.ResultAt(index).CrackSummaryMeta)
    data(rowIndex, COL_STATUS) = batch.ResultAt(index).CrackSummaryStatus

    If CrackWidthStatusText(batch, index, policy) = policy.NotApplicable And _
            LongitudinalStatusText(batch, index, policy) = policy.NotApplicable And _
            Not HasCrackFormationData(batch, index) And _
            Not batch.ResultAt(index).CrackCurrentStateMeta.Applies And _
            Not batch.ResultAt(index).CrackFormationMeta.Applies Then
        Exit Sub
    End If

    data(rowIndex, COL_REF_N) = OutputForce(batch.CrackReferenceN(index), units)
    data(rowIndex, COL_REF_MX) = OutputMx(batch.CrackReferenceMx(index), units)
    data(rowIndex, COL_REF_MY) = OutputMy(batch.CrackReferenceMy(index), units)
    data(rowIndex, COL_LOAD_N) = OutputForce(batch.CrackLoadPointN(index), units)
    data(rowIndex, COL_LOAD_MX) = OutputMx(batch.CrackLoadPointMx(index), units)
    data(rowIndex, COL_LOAD_MY) = OutputMy(batch.CrackLoadPointMy(index), units)

    PutFormationData data, rowIndex, batch, index, units, policy
    PutCrackedStateData data, rowIndex, batch, index, policy
    PutCrackWidthData data, rowIndex, batch, index, units, policy
    PutLongitudinalData data, rowIndex, batch, index, units, policy
End Sub

' Возвращает позицию LC внутри data-части подробной таблицы.
' Основной поток читает offset из rngLoadCombinations; прямые тесты без offset продолжают выводиться компактно.
Private Function OutputDataOffset(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal dataRows As Long) As Long
    OutputDataOffset = batch.SourceDataOffset(index)
    If OutputDataOffset < 1 Then OutputDataOffset = index
    If OutputDataOffset > dataRows Then OutputDataOffset = index
    If OutputDataOffset > dataRows Then OutputDataOffset = dataRows
End Function

' Выводит доступные результаты образования трещины и статусы Pre/Post НДС.
' При отсутствии точки не создает нулевые Ncrc/Mcrc; общий исход берется из
' Formation result, а не восстанавливается из чисел writer-ом.
Private Sub PutFormationData(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal units As CUnitSystem, ByVal policy As CResultStatusPolicy)
    Dim hasFormationData As Boolean
    hasFormationData = HasCrackFormationData(batch, index)

    If CrackFormationStatusText(batch, index, policy) = policy.NotApplicable And Not hasFormationData Then
        data(rowIndex, COL_FORMATION_STATE) = policy.NotApplicable
        Exit Sub
    End If

    If Len(batch.ResultAt(index).CrackResult.Formation.FormationMethod) > 0 Then _
        data(rowIndex, COL_FORMATION_METHOD) = batch.ResultAt(index).CrackResult.Formation.FormationMethod

    If batch.ResultAt(index).CrackResult.Formation.Ncrc > SMALL_VALUE Then
        data(rowIndex, COL_ARED) = OutputArea(batch.ResultAt(index).CrackResult.Formation.Ared, units)
        data(rowIndex, COL_RBT_SER) = OutputStress(batch.ResultAt(index).CrackResult.Formation.RbtSer, units)
        data(rowIndex, COL_NCRC) = OutputForce(batch.ResultAt(index).CrackResult.Formation.Ncrc, units)
        data(rowIndex, COL_LAMBDA_NCRC) = batch.ResultAt(index).CrackResult.Formation.LambdaCrc
    Else
        data(rowIndex, COL_NCRC) = policy.NotApplicable
    End If

    If (Not batch.ResultAt(index).CrackResult.Width.CentralTensionBranch) And hasFormationData Then
        data(rowIndex, COL_FORMATION_NCRC) = OutputForce(batch.ResultAt(index).CrackResult.Formation.FormationNcrc, units)
    ElseIf Not batch.ResultAt(index).CrackResult.Width.CentralTensionBranch And batch.ResultAt(index).CrackResult.Formation.Mcrc > SMALL_VALUE Then
        data(rowIndex, COL_FORMATION_NCRC) = policy.NotApplicable
    End If

    If batch.ResultAt(index).CrackResult.Formation.Mcrc > SMALL_VALUE Then
        data(rowIndex, COL_MCRC) = OutputMomentMagnitude(batch.ResultAt(index).CrackResult.Formation.Mcrc, units)
        data(rowIndex, COL_LAMBDA_MCRC) = batch.ResultAt(index).CrackResult.Formation.LambdaCrc
    Else
        data(rowIndex, COL_MCRC) = policy.NotApplicable
    End If

    data(rowIndex, COL_PRE_CRACK_STATUS) = NamedStateStatus(batch, index, sstPreCrackState, policy)
    data(rowIndex, COL_POST_CRACK_STATUS) = NamedStateStatus(batch, index, sstPostCrackState, policy)
    PutNamedStateStrains data, rowIndex, batch.ResultAt(index).CrackResult.Formation.PreCrackState, COL_PRE_EBC, COL_PRE_EBT, True
    PutNamedStateStrains data, rowIndex, batch.ResultAt(index).CrackResult.Formation.PostCrackState, COL_POST_EBC, COL_POST_EBT, False

    If batch.ResultAt(index).CrackResult.Formation.ConfirmedNotCracked Then
        data(rowIndex, COL_FORMATION_STATE) = "NotCracked"
    ElseIf hasFormationData Then
        data(rowIndex, COL_FORMATION_STATE) = IIf(batch.ResultAt(index).CrackResult.Formation.CrackFormed, "Cracked", "NotCracked")
    Else
        data(rowIndex, COL_FORMATION_STATE) = CrackFormationStatusText(batch, index, policy)
    End If
End Sub

' Возвращает статус сохраненного named-state. Если состояние штатно не создавалось,
' detailed-таблица показывает N/A, а не общий статус всей проверки трещин.
Private Function NamedStateStatus(ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal stateType As ESectionStateType, ByVal policy As CResultStatusPolicy) As String
    Dim stateResult As CSectionStateResult
    Set stateResult = batch.ResultAt(index).StateRepository.FindState(stateType)
    If stateResult Is Nothing Then
        NamedStateStatus = policy.NotApplicable
    Else
        NamedStateStatus = policy.ExternalStatus(stateResult.ResultMeta)
    End If
End Function

' Выводит экстремумы только полученного состояния. Pre всегда показывает
' фактическое eps_bt; остальные - при активном растяжении бетона своей spec.
' Предельные деформации и нули вместо отсутствующего НДС не подставляются.
Private Sub PutNamedStateStrains(ByRef data() As Variant, ByVal rowIndex As Long, ByVal state As CSectionStateResult, _
        ByVal compressionColumn As Long, ByVal tensionColumn As Long, ByVal beforeCrack As Boolean)
    If state Is Nothing Then Exit Sub
    If Not state.Converged Then Exit Sub
    If state.HasConcreteCompression Then data(rowIndex, compressionColumn) = state.MinConcreteStrain
    If beforeCrack Or state.MaterialSpec.ConcreteTension = ctkUseDiagram Then data(rowIndex, tensionColumn) = state.MaxConcreteStrain
End Sub

' Определяет наличие ненулевых сохраненных усилий пороговой точки для вывода.
' Это доступность числовых полей, не отдельная проверка сходимости или физики.
Private Function HasCrackFormationData(ByVal batch As CBatchSectionCalculator, ByVal index As Long) As Boolean
    HasCrackFormationData = batch.ResultAt(index).CrackResult.Formation.Ncrc > SMALL_VALUE Or _
        Abs(batch.ResultAt(index).CrackResult.Formation.FormationNcrc) > SMALL_VALUE Or _
        batch.ResultAt(index).CrackResult.Formation.Mcrc > SMALL_VALUE
End Function

' Записывает контроль пригодности текущего CrackedState для дальнейшего
' расчета трещин. Это не проверка факта образования нормальной трещины:
' трещинообразование уже проверено отдельным блоком Mcrc/Ncrc. Здесь
' фиксируется, удалось ли найти равновесное НДС полного заданного сочетания
' по модели раскрытой трещины и осталось ли это состояние в физической части
' диаграмм материалов. Если равновесие найдено только через продолжение
' диаграммы или solver не сошелся, значения ls, sigma_s, a_crc и проверка
' продольных трещин не рассчитываются, потому что опираться на такое НДС
' как на физическое состояние сечения нельзя.
Private Sub PutCrackedStateData(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal policy As CResultStatusPolicy)
    Dim stateResult As CSectionStateResult
    Set stateResult = batch.ResultAt(index).StateRepository.FindState(sstCrackedState)

    If stateResult Is Nothing Then
        If CrackCurrentStateStatusText(batch, index, policy) <> policy.NotApplicable Then _
            data(rowIndex, COL_CRACK_STATE_STATUS) = CrackCurrentStateStatusText(batch, index, policy)
        Exit Sub
    End If

    data(rowIndex, COL_CRACK_STATE_DIAGEXT) = ExtensionUsedText(batch, index, stateResult)
    data(rowIndex, COL_CRACK_STATE_STATUS) = CrackCurrentStateStatusText(batch, index, policy)
    PutNamedStateStrains data, rowIndex, stateResult, COL_CURRENT_EBC, COL_CURRENT_EBT, False
End Sub

' Показывает, использовало ли текущее CrackedState продолжение диаграммы.
' Значение "-" означает, что General.DiagramExtension был выключен
' для этого LC, поэтому extension в принципе не мог применяться.
Private Function ExtensionUsedText(ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal stateResult As CSectionStateResult) As String
    If Not batch.ResultAt(index).DiagramExtensionEnabled Then
        ExtensionUsedText = "-"
    ElseIf stateResult.ExtensionUsed Then
        ExtensionUsedText = "yes"
    Else
        ExtensionUsedText = "no"
    End If
End Function

' Записывает уже рассчитанную ширину и подготовленные величины ее формулы.
' Успех Formation не является условием вывода: допустимый fallback psi_s = 1
' сохраняется в самостоятельном Width-result. Writer не повторяет проверку.
Private Sub PutCrackWidthData(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal units As CUnitSystem, ByVal policy As CResultStatusPolicy)
    Dim width As CCrackWidthResult
    Set width = batch.ResultAt(index).CrackResult.Width
    data(rowIndex, COL_SP35_STATUS) = policy.NotApplicable
    data(rowIndex, COL_ACRC_STATUS) = policy.NotApplicable
    If width.StandardCode = "SP35" Then
        data(rowIndex, COL_SP35_STATUS) = policy.ExternalStatus(width.ResultMeta)
        If width.ResultMeta.Calculated Then PutSP35WidthData data, rowIndex, width, units
        Exit Sub
    End If

    If Not batch.ResultAt(index).CrackWidthMeta.Calculated Then
        data(rowIndex, COL_ACRC_STATUS) = CrackWidthStatusText(batch, index, policy)
        Exit Sub
    End If

    If Not policy.IsFinished(CrackWidthStatusText(batch, index, policy)) Then
        data(rowIndex, COL_ACRC_STATUS) = CrackWidthStatusText(batch, index, policy)
        Exit Sub
    End If

    If Not batch.ResultAt(index).CrackResult.Width.CentralTensionBranch Then
        data(rowIndex, COL_XT) = OutputLength(batch.ResultAt(index).CrackResult.Width.TensionDepth, units)
        data(rowIndex, COL_H) = OutputLength(batch.ResultAt(index).CrackResult.Width.SectionDepthH, units)
        data(rowIndex, COL_AS_COVER) = OutputLength(batch.ResultAt(index).CrackResult.Width.CoverA, units)
        data(rowIndex, COL_HBT) = OutputLength(batch.ResultAt(index).CrackResult.Width.EffectiveZoneDepth, units)
    End If
    data(rowIndex, COL_ABT) = OutputArea(batch.ResultAt(index).CrackResult.Width.Abt, units)
    data(rowIndex, COL_AS) = OutputArea(batch.ResultAt(index).CrackResult.Width.AsTension, units)
    data(rowIndex, COL_DS_EQ) = OutputLength(batch.ResultAt(index).CrackResult.Width.DsEquivalent, units)
    data(rowIndex, COL_LS0) = OutputLength(batch.ResultAt(index).CrackResult.Width.CrackSpacingRaw, units)
    data(rowIndex, COL_LS) = OutputLength(batch.ResultAt(index).CrackResult.Width.CrackSpacing, units)
    data(rowIndex, COL_PHI1) = batch.ResultAt(index).CrackResult.Width.Phi1
    data(rowIndex, COL_PHI2) = batch.ResultAt(index).CrackResult.Width.Phi2
    data(rowIndex, COL_PHI3) = batch.ResultAt(index).CrackResult.Width.Phi3
    data(rowIndex, COL_PSI_S) = batch.ResultAt(index).CrackResult.Width.PsiS
    data(rowIndex, COL_REBARS) = batch.ResultAt(index).CrackResult.Width.TensionRebarIds
    data(rowIndex, COL_NREBARS) = batch.ResultAt(index).CrackResult.Width.TensionRebarCount
    data(rowIndex, COL_SIGMA_S_CRC) = OptionalStress(batch.ResultAt(index).CrackResult.Width.SigmaSCrc, units, policy)
    data(rowIndex, COL_SIGMA_S) = OutputStress(batch.ResultAt(index).CrackResult.Width.SigmaS, units)
    data(rowIndex, COL_ES) = OutputStress(batch.ResultAt(index).CrackResult.Width.SteelEs, units)
    data(rowIndex, COL_ACRC) = OutputLength(batch.ResultAt(index).CrackResult.Width.CrackWidth, units)
    data(rowIndex, COL_ACRC_ULT) = OutputLength(batch.ResultAt(index).CrackResult.Width.AllowableCrackWidth, units)
    If batch.ResultAt(index).CrackResult.Width.ReserveFactor > SMALL_VALUE Then _
        data(rowIndex, COL_ACRC_RATIO) = batch.ResultAt(index).CrackResult.Width.ReserveFactor
    data(rowIndex, COL_ACRC_STATUS) = CrackWidthStatusText(batch, index, policy)
End Sub

' Записывает проверку продольных трещин по сжатому бетону.
Private Sub PutLongitudinalData(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal units As CUnitSystem, ByVal policy As CResultStatusPolicy)
    If Not policy.IsFinished(LongitudinalStatusText(batch, index, policy)) Then
        data(rowIndex, COL_LONG_STATUS) = LongitudinalStatusText(batch, index, policy)
        Exit Sub
    End If

    data(rowIndex, COL_SIGMA_B_MAX) = OutputStress(batch.ResultAt(index).CrackResult.Longitudinal.MaxCompressionStress, units)
    data(rowIndex, COL_RB_MC2) = OutputStress(batch.ResultAt(index).CrackResult.Longitudinal.RbMc2, units)
    If batch.ResultAt(index).CrackResult.Longitudinal.ReserveFactor > SMALL_VALUE Then _
        data(rowIndex, COL_LONG_RATIO) = batch.ResultAt(index).CrackResult.Longitudinal.ReserveFactor
    data(rowIndex, COL_LONG_STATUS) = LongitudinalStatusText(batch, index, policy)
End Sub

' Оформляет шапки, границы, числовые форматы и подсветку статусов.
Private Sub FormatTable(ByVal target As Object, ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    target.Font.Name = "Arial"
    target.Font.Size = 9
    target.HorizontalAlignment = -4108
    target.VerticalAlignment = -4108
    target.WrapText = False
    target.Rows.RowHeight = 15

    FormatBlock target, COL_LEFT_FIRST, COL_LEFT_LAST
    FormatBlock target, COL_MAIN_FIRST, COL_CRACK_STATE_STATUS
    FormatBlock target, COL_SP35_AR, COL_SP35_STATUS
    FormatBlock target, COL_XT, COL_ACRC_STATUS
    FormatBlock target, COL_SIGMA_B_MAX, COL_LONG_STATUS
    target.Columns(COL_SEPARATOR).Clear
    target.Columns(COL_FORMATION_SEPARATOR).Clear
    target.Columns(COL_SP35_SEPARATOR).Clear
    target.Columns(COL_SP63_SEPARATOR).Clear

    target.Resize(HEADER_ROWS, TOTAL_COLUMNS).HorizontalAlignment = -4108
    target.Resize(HEADER_ROWS, TOTAL_COLUMNS).VerticalAlignment = -4108
    target.Resize(HEADER_ROWS, TOTAL_COLUMNS).WrapText = True
    target.Rows(HEADER_NOTES_ROW).Font.Bold = False
    target.Rows(HEADER_NOTES_ROW).HorizontalAlignment = -4131
    target.Rows(HEADER_NOTES_ROW).WrapText = False
    target.Rows(HEADER_LABEL_ROW).WrapText = False
    target.Cells(HEADER_LABEL_ROW, COL_RBT_SER).WrapText = False
    ' Вертикально объединенные заголовки остаются многострочными, как в ТЗ.
    ' Пояснения и обозначения остальных столбцов не получают перенос.
    Dim headerColumn As Variant
    For Each headerColumn In Array(COL_ID, COL_RESULT_COMMENT, COL_STATUS)
        target.Cells(1, CLng(headerColumn)).WrapText = True
    Next headerColumn

    Dim dataRows As Long
    dataRows = target.Rows.Count - HEADER_ROWS
    If dataRows > 0 Then
        target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).HorizontalAlignment = -4108
        target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).VerticalAlignment = -4108
        target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).NumberFormat = "0.000"
        target.Offset(HEADER_ROWS, COL_ID - 1).Resize(dataRows, 3).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_RESULT_COMMENT - 1).Resize(dataRows, 1).HorizontalAlignment = -4131
        target.Offset(HEADER_ROWS, COL_REBARS - 1).Resize(dataRows, 1).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_PRE_CRACK_STATUS - 1).Resize(dataRows, 1).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_POST_CRACK_STATUS - 1).Resize(dataRows, 1).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_FORMATION_STATE - 1).Resize(dataRows, 1).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_SP35_BETA - 1).Resize(dataRows, 3).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_SP35_REBARS - 1).Resize(dataRows, 1).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_SP35_BETA - 1).Resize(dataRows, 3).WrapText = False
        target.Offset(HEADER_ROWS, COL_SP35_REBARS - 1).Resize(dataRows, 1).WrapText = False
        target.Offset(HEADER_ROWS, COL_SP35_STATUS - 1).Resize(dataRows, 1).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_PRE_EBC - 1).Resize(dataRows, 2).NumberFormat = "0.000000"
        target.Offset(HEADER_ROWS, COL_POST_EBC - 1).Resize(dataRows, 2).NumberFormat = "0.000000"
        target.Offset(HEADER_ROWS, COL_CURRENT_EBC - 1).Resize(dataRows, 2).NumberFormat = "0.000000"
        target.Offset(HEADER_ROWS, COL_CRACK_STATE_DIAGEXT - 1).Resize(dataRows, 2).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_ACRC_STATUS - 1).Resize(dataRows, 1).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_LONG_STATUS - 1).Resize(dataRows, 1).NumberFormat = "@"
        ' Сохраняем одинаковую физическую точность 0.001 мм: ширина в метрах
        ' не должна выглядеть нулевой из-за общего трехзначного формата таблицы.
        Dim widthFormat As String
        widthFormat = CrackWidthNumberFormat(units)
        target.Offset(HEADER_ROWS, COL_SP35_ACRC - 1).Resize(dataRows, 2).NumberFormat = widthFormat
        target.Offset(HEADER_ROWS, COL_ACRC - 1).Resize(dataRows, 2).NumberFormat = widthFormat
        ApplyStatusCells target, batch, dataRows
    End If
    ApplyHeaderSubscripts target
End Sub

' Возвращает точность отображения ширины/допуска в текущей OUTPUT Length.
' Форматирование не округляет сохраненное число и не меняет размер столбца.
Private Function CrackWidthNumberFormat(ByVal units As CUnitSystem) As String
    Select Case OutputLengthUnit(units)
        Case "m": CrackWidthNumberFormat = "0.000000"
        Case "cm": CrackWidthNumberFormat = "0.0000"
        Case Else: CrackWidthNumberFormat = "0.000"
    End Select
End Function

' Оформляет один непрерывный блок таблицы, не затрагивая пустые разделители.
Private Sub FormatBlock(ByVal target As Object, ByVal firstCol As Long, ByVal lastCol As Long)
    Dim block As Object
    Set block = target.Offset(0, firstCol - 1).Resize(target.Rows.Count, lastCol - firstCol + 1)
    block.Borders.LineStyle = 1
    block.Borders.Weight = 2
    block.Rows(1).Resize(HEADER_ROWS).Font.Bold = True
    target.Offset(0, firstCol - 1).Resize(HEADER_ROWS, lastCol - firstCol + 1).Interior.Color = HEADER_COLOR
End Sub

' Красит блоки строки по статусу общей проверки, нормальных и продольных трещин.
Private Sub ApplyStatusCells(ByVal target As Object, ByVal batch As CBatchSectionCalculator, ByVal dataRows As Long)
    If batch Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To batch.Count
        Dim rowOffset As Long
        rowOffset = OutputDataOffset(batch, i, dataRows)
        ApplyStatusCell target, HEADER_ROWS + rowOffset, COL_STATUS
        ApplyStatusCell target, HEADER_ROWS + rowOffset, COL_PRE_CRACK_STATUS
        ApplyStatusCell target, HEADER_ROWS + rowOffset, COL_POST_CRACK_STATUS
        ApplyFormationStatusCell target, HEADER_ROWS + rowOffset
        ApplyStatusCell target, HEADER_ROWS + rowOffset, COL_CRACK_STATE_STATUS
        ApplyStatusCell target, HEADER_ROWS + rowOffset, COL_ACRC_STATUS
        ApplyStatusCell target, HEADER_ROWS + rowOffset, COL_SP35_STATUS
        ApplyStatusCell target, HEADER_ROWS + rowOffset, COL_LONG_STATUS
    Next i
End Sub

' Пишет снимок выполненной формулы СП 35. Критический якорь, beta/psi и область
' уже сохранены calculator-ом; writer лишь переводит единицы и форматирует.
Private Sub PutSP35WidthData(ByRef data() As Variant, ByVal rowIndex As Long, ByVal width As CCrackWidthResult, ByVal units As CUnitSystem)
    data(rowIndex, COL_SP35_AR) = OutputArea(width.Abt, units)
    data(rowIndex, COL_SP35_BETA) = SP35RowsText(width, "Beta", units)
    data(rowIndex, COL_SP35_N) = SP35RowsText(width, "Count", units)
    data(rowIndex, COL_SP35_D) = SP35RowsText(width, "Diameter", units)
    data(rowIndex, COL_SP35_RR) = OutputLength(width.SP35ReinforcementRadius, units)
    data(rowIndex, COL_SP35_PSI) = OutputLength(width.SP35PsiCm * 10#, units)
    data(rowIndex, COL_SP35_REBARS) = SP35RowsText(width, "Rebars", units)
    data(rowIndex, COL_SP35_NREBARS) = width.TensionRebarCount
    data(rowIndex, COL_SP35_SIGMA) = OutputStress(width.SigmaS, units)
    data(rowIndex, COL_SP35_ES) = OutputStress(width.SteelEs, units)
    data(rowIndex, COL_SP35_ACRC) = OutputLength(width.CrackWidth, units)
    data(rowIndex, COL_SP35_ALLOWABLE) = OutputLength(width.AllowableCrackWidth, units)
    If width.ReserveFactor > SMALL_VALUE Then data(rowIndex, COL_SP35_RATIO) = width.ReserveFactor
End Sub

' Форматирует все ряды критической области. Несколько групп одного ряда
' сохраняют соответствие beta/n/d/стержней по G-id. Диаметры реальные, не d_ref.
Private Function SP35RowsText(ByVal width As CCrackWidthResult, ByVal field As String, ByVal units As CUnitSystem) As String
    Dim prepared As CSP35CrackData, rows As Variant, members As Variant, row As Long, i As Long, j As Long
    Dim group As Long, count As Long, segment As String, value As String
    Set prepared = width.SP35Data: rows = prepared.CandidateRows(width.SP35CriticalCandidate)
    For row = 1 To prepared.CandidateRowCount(width.SP35CriticalCandidate)
        count = 0: segment = vbNullString
        For i = LBound(rows, 1) To UBound(rows, 1)
            If CLng(rows(i, 2)) = row Then count = count + 1
        Next i
        For i = LBound(rows, 1) To UBound(rows, 1)
            If CLng(rows(i, 2)) = row Then
                group = CLng(rows(i, 1)): value = vbNullString
                Select Case field
                    Case "Beta": value = Replace$(Format$(prepared.CandidateGroupBeta(width.SP35CriticalCandidate, group), "0.00"), ",", ".")
                    Case "Count": value = CStr(prepared.CandidateGroupBarCount(width.SP35CriticalCandidate, group))
                    Case Else
                        members = prepared.CandidateGroupMembers(width.SP35CriticalCandidate, group)
                        For j = LBound(members) To UBound(members)
                            If Len(value) > 0 Then value = value & ", "
                            If field = "Diameter" Then
                                value = value & FormatRowNumber(OutputLength(prepared.BarDiameter(CLng(members(j))), units))
                            Else
                                value = value & prepared.BarID(CLng(members(j)))
                            End If
                        Next j
                End Select
                If count > 1 Then value = "G" & CStr(group) & "=" & value
                If Len(segment) > 0 Then segment = segment & "; "
                segment = segment & value
            End If
        Next i
        If Len(SP35RowsText) > 0 Then SP35RowsText = SP35RowsText & " | "
        SP35RowsText = SP35RowsText & CStr(row) & " ряд: " & segment
    Next row
End Function

' Убирает пустой десятичный знак в текстовой расшифровке диаметра.
Private Function FormatRowNumber(ByVal value As Double) As String
    FormatRowNumber = Replace$(Format$(value, "0.########"), ",", ".")
    If Right$(FormatRowNumber, 1) = "." Then FormatRowNumber = Left$(FormatRowNumber, Len(FormatRowNumber) - 1)
End Function

' Применяет общую палитру к исходу поиска трещинообразования. Cracked и
' NotCracked описывают наличие трещины, а не успешность проверки; они не
' передаются в нормализатор внешних статусов и остаются без статусной заливки.
Private Sub ApplyFormationStatusCell(ByVal target As Object, ByVal rowIndex As Long)
    Dim text As String
    text = CStr(target.Cells(rowIndex, COL_FORMATION_STATE).Value2)
    If text = "Cracked" Or text = "NotCracked" Then Exit Sub
    ApplyStatusCell target, rowIndex, COL_FORMATION_STATE
End Sub

' Красит одну статусную ячейку detailed-блока, оставляя числовые поля без статусной заливки.
Private Sub ApplyStatusCell(ByVal target As Object, ByVal rowIndex As Long, ByVal columnIndex As Long)
    Dim cell As Object
    Set cell = target.Cells(rowIndex, columnIndex)
    ApplyStatusFill cell, CStr(cell.Value2)
End Sub

' Выставляет нижние индексы в обозначениях нижнего яруса заголовков.
Private Sub ApplyHeaderSubscripts(ByVal target As Object)
    Dim labels As Object
    Set labels = target.Rows(HEADER_LABEL_ROW)
    Dim column As Variant
    For Each column In Array(COL_PRE_EBC, COL_POST_EBC, COL_CURRENT_EBC): ApplySubscriptToken labels.Cells(1, CLng(column)), "bc": Next column
    For Each column In Array(COL_PRE_EBT, COL_POST_EBT, COL_CURRENT_EBT): ApplySubscriptToken labels.Cells(1, CLng(column)), "bt": Next column
    ApplySubscriptToken labels.Cells(1, COL_SP35_AR), "r"
    ApplySubscriptToken labels.Cells(1, COL_SP35_RR), "r"
    ApplySubscriptToken labels.Cells(1, COL_SP35_SIGMA), "s"
    ApplySubscriptToken labels.Cells(1, COL_SP35_ES), "s"
    ApplySubscriptToken labels.Cells(1, COL_SP35_NREBARS), "rebars"
    For Each column In Array(COL_SP35_ACRC, COL_SP35_ALLOWABLE, COL_SP35_RATIO): ApplySubscriptToken labels.Cells(1, CLng(column)), "crc": Next column
    ApplySubscriptToken labels.Cells(1, COL_SP35_ALLOWABLE), "ult"
    ApplySubscriptToken labels.Cells(1, COL_SP35_RATIO), "ult"

    ApplySubscriptToken target.Cells(2, COL_DS_EQ), "s"

    ApplySubscriptToken labels.Cells(1, COL_ARED), "red"
    ApplySubscriptToken labels.Cells(1, COL_RBT_SER), "bt"
    ApplySubscriptToken labels.Cells(1, COL_RBT_SER), "ser"
    ApplySubscriptToken labels.Cells(1, COL_NCRC), "crc"
    ApplySubscriptToken labels.Cells(1, COL_LAMBDA_NCRC), "crc"
    ApplySubscriptToken labels.Cells(1, COL_FORMATION_NCRC), "crc"
    ApplySubscriptToken labels.Cells(1, COL_MCRC), "crc"
    ApplySubscriptToken labels.Cells(1, COL_LAMBDA_MCRC), "crc"
    ApplySubscriptToken labels.Cells(1, COL_XT), "t"
    ApplySubscriptToken labels.Cells(1, COL_AS_COVER), "s"
    ApplySubscriptToken labels.Cells(1, COL_HBT), "bt"
    ApplySubscriptToken labels.Cells(1, COL_ABT), "bt"
    ApplySubscriptToken labels.Cells(1, COL_AS), "s"
    ApplySubscriptToken labels.Cells(1, COL_DS_EQ), "s"
    ApplySubscriptToken labels.Cells(1, COL_DS_EQ), "eq"
    ApplySubscriptToken labels.Cells(1, COL_LS0), "s"
    ApplySubscriptToken labels.Cells(1, COL_LS0), "0"
    ApplySubscriptToken labels.Cells(1, COL_LS), "s"
    ApplySubscriptToken labels.Cells(1, COL_PHI1), "1"
    ApplySubscriptToken labels.Cells(1, COL_PHI2), "2"
    ApplySubscriptToken labels.Cells(1, COL_PHI3), "3"
    ApplySubscriptToken labels.Cells(1, COL_PSI_S), "s"
    ApplySubscriptToken labels.Cells(1, COL_NREBARS), "rebars"
    ApplySubscriptToken labels.Cells(1, COL_SIGMA_S_CRC), "s"
    ApplySubscriptToken labels.Cells(1, COL_SIGMA_S_CRC), "crc"
    ApplySubscriptToken labels.Cells(1, COL_SIGMA_S), "s"
    ApplySubscriptToken labels.Cells(1, COL_ES), "s"
    ApplySubscriptToken labels.Cells(1, COL_ACRC), "crc"
    ApplySubscriptToken labels.Cells(1, COL_ACRC_ULT), "crc"
    ApplySubscriptToken labels.Cells(1, COL_ACRC_ULT), "ult"
    ApplySubscriptToken labels.Cells(1, COL_ACRC_RATIO), "crc"
    ApplySubscriptToken labels.Cells(1, COL_ACRC_RATIO), "ult"
    ApplySubscriptToken labels.Cells(1, COL_SIGMA_B_MAX), "b"
    ApplySubscriptToken labels.Cells(1, COL_SIGMA_B_MAX), "max"
    ApplySubscriptToken labels.Cells(1, COL_RB_MC2), "b"
    ApplySubscriptToken labels.Cells(1, COL_RB_MC2), "mc2"
    ApplySubscriptToken labels.Cells(1, COL_LONG_RATIO), "b"
    ApplySubscriptToken labels.Cells(1, COL_LONG_RATIO), "mc2"
End Sub

' Ставит нижний индекс найденному текстовому токену внутри одной ячейки.
Private Sub ApplySubscriptToken(ByVal cell As Object, ByVal token As String)
    Dim textValue As String
    textValue = CStr(cell.Value2)
    Dim startPos As Long
    startPos = InStr(1, textValue, token, vbBinaryCompare)
    Do While startPos > 0
        cell.Characters(startPos, Len(token)).Font.Subscript = True
        startPos = InStr(startPos + Len(token), textValue, token, vbBinaryCompare)
    Loop
End Sub

' Возвращает внешний статус поиска образования нормальной трещины.
Private Function CrackFormationStatusText(ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal policy As CResultStatusPolicy) As String
    Dim meta As CResultMeta
    Set meta = batch.ResultAt(index).CrackFormationMeta
    CrackFormationStatusText = policy.ExternalStatus(meta)
End Function

' Возвращает внешний статус прямого CrackedState от заданного сочетания.
Private Function CrackCurrentStateStatusText(ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal policy As CResultStatusPolicy) As String
    Dim meta As CResultMeta
    Set meta = batch.ResultAt(index).CrackCurrentStateMeta
    CrackCurrentStateStatusText = policy.ExternalStatus(meta)
End Function

' Возвращает внешний статус формульной проверки ширины нормальной трещины.
Private Function CrackWidthStatusText(ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal policy As CResultStatusPolicy) As String
    Dim meta As CResultMeta
    Set meta = batch.ResultAt(index).CrackWidthMeta
    CrackWidthStatusText = policy.ExternalStatus(meta)
End Function

' Возвращает внешний статус проверки продольных трещин.
Private Function LongitudinalStatusText(ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal policy As CResultStatusPolicy) As String
    Dim meta As CResultMeta
    Set meta = batch.ResultAt(index).LongitudinalCrackMeta
    LongitudinalStatusText = policy.ExternalStatus(meta)
End Function

' Возвращает готовое пользовательское пояснение из ResultMeta блока.
' Writer не анализирует технические StopReason и не собирает комментарий заново.
Private Function MetaComment(ByVal meta As CResultMeta) As String
    If meta Is Nothing Then Exit Function
    MetaComment = meta.ResultComment
End Function

' Возвращает напряжение или N/A для величин, которые расчетно могут отсутствовать.
Private Function OptionalStress(ByVal value As Double, ByVal units As CUnitSystem, ByVal policy As CResultStatusPolicy) As Variant
    If Abs(value) > SMALL_VALUE Then
        OptionalStress = OutputStress(value, units)
    Else
        OptionalStress = policy.NotApplicable
    End If
End Function

' Объединяет прямоугольный участок шапки.
Private Sub MergeRange(ByVal target As Object, ByVal firstRow As Long, ByVal firstCol As Long, _
        ByVal lastRow As Long, ByVal lastCol As Long)
    If lastRow < firstRow Or lastCol < firstCol Then Exit Sub
    target.Worksheet.Range(target.Cells(firstRow, firstCol), target.Cells(lastRow, lastCol)).Merge
End Sub

' Ниже собраны маленькие адаптеры единиц для crack-блока: они оставляют внутренние единицы,
' если CUnitSystem не передан, и применяют пользовательские выходные единицы в штатном расчете.
Private Function OutputLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputLength = value Else OutputLength = units.InternalLengthToOutput(value)
End Function

' Возвращает подпись выходной единицы длины для заголовков crack-таблицы.
Private Function OutputLengthUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputLengthUnit = "mm" Else OutputLengthUnit = units.OutputLengthUnit
End Function

' Переводит площадь из внутренних мм2 в выбранную пользователем единицу площади.
Private Function OutputArea(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputArea = value Else OutputArea = units.InternalAreaToOutput(value)
End Function

' Возвращает подпись выходной единицы площади для заголовков crack-таблицы.
Private Function OutputAreaUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputAreaUnit = "mm2" Else OutputAreaUnit = units.OutputAreaUnit
End Function

' Переводит продольную силу из внутренних Н в пользовательскую выходную систему знаков и единиц.
Private Function OutputForce(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputForce = value Else OutputForce = units.InternalForceToOutput(value)
End Function

' Возвращает подпись выходной единицы силы для заголовков crack-таблицы.
Private Function OutputForceUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputForceUnit = "N" Else OutputForceUnit = units.OutputForceUnit
End Function

' Переводит момент Mx с учетом пользовательской системы знаков.
Private Function OutputMx(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMx = value Else OutputMx = units.InternalMomentMxToOutput(value)
End Function

' Переводит момент My с учетом пользовательской системы знаков.
Private Function OutputMy(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMy = value Else OutputMy = units.InternalMomentMyToOutput(value)
End Function

' Переводит модуль момента; знак здесь не применяется, потому что величина скалярная.
Private Function OutputMomentMagnitude(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMomentMagnitude = value Else OutputMomentMagnitude = units.InternalMomentMagnitudeToOutput(value)
End Function

' Возвращает подпись выходной единицы момента для заголовков crack-таблицы.
Private Function OutputMomentUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputMomentUnit = "N*mm" Else OutputMomentUnit = units.OutputMomentUnit
End Function

' Переводит напряжение из внутренних МПа в пользовательскую выходную единицу.
Private Function OutputStress(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputStress = value Else OutputStress = units.InternalStressToOutput(value)
End Function

' Возвращает подпись выходной единицы напряжения для заголовков crack-таблицы.
Private Function OutputStressUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputStressUnit = "MPa" Else OutputStressUnit = units.OutputStressUnit
End Function

' Возвращает единый словарь статусов для подробного crack-блока.
Private Function StatusPolicy() As CResultStatusPolicy
    Static policy As CResultStatusPolicy
    If policy Is Nothing Then Set policy = New CResultStatusPolicy
    Set StatusPolicy = policy
End Function
