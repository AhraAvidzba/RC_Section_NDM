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
Private Const TOTAL_COLUMNS As Long = 43 ' Полная ширина таблицы, включая пустой разделитель.

Private Const COL_LEFT_FIRST As Long = 1
Private Const COL_LEFT_LAST As Long = 8
Private Const COL_SEPARATOR As Long = 9
Private Const COL_MAIN_FIRST As Long = 10
Private Const COL_MAIN_LAST As Long = 43

Private Const COL_ID As Long = 1
Private Const COL_STATUS As Long = 2
Private Const COL_REF_N As Long = 3
Private Const COL_REF_MX As Long = 4
Private Const COL_REF_MY As Long = 5
Private Const COL_LOAD_N As Long = 6
Private Const COL_LOAD_MX As Long = 7
Private Const COL_LOAD_MY As Long = 8

Private Const COL_ARED As Long = 10
Private Const COL_RBT_SER As Long = 11
Private Const COL_NCRC As Long = 12
Private Const COL_LAMBDA_NCRC As Long = 13
Private Const COL_FORMATION_METHOD As Long = 14
Private Const COL_FORMATION_NCRC As Long = 15
Private Const COL_MCRC As Long = 16
Private Const COL_LAMBDA_MCRC As Long = 17
Private Const COL_FORMATION_STATE As Long = 18
Private Const COL_XT As Long = 19
Private Const COL_H As Long = 20
Private Const COL_AS_COVER As Long = 21
Private Const COL_HBT As Long = 22
Private Const COL_ABT As Long = 23
Private Const COL_AS As Long = 24
Private Const COL_DS_EQ As Long = 25
Private Const COL_LS0 As Long = 26
Private Const COL_LS As Long = 27
Private Const COL_PHI1 As Long = 28
Private Const COL_PHI2 As Long = 29
Private Const COL_PHI3 As Long = 30
Private Const COL_PSI_S As Long = 31
Private Const COL_REBARS As Long = 32
Private Const COL_NREBARS As Long = 33
Private Const COL_SIGMA_S_CRC As Long = 34
Private Const COL_SIGMA_S As Long = 35
Private Const COL_ACRC As Long = 36
Private Const COL_ACRC_ULT As Long = 37
Private Const COL_ACRC_RATIO As Long = 38
Private Const COL_ACRC_STATUS As Long = 39
Private Const COL_SIGMA_B_MAX As Long = 40
Private Const COL_RB_MC2 As Long = 41
Private Const COL_LONG_RATIO As Long = 42
Private Const COL_LONG_STATUS As Long = 43

Private Const HEADER_COLOR As Long = &HD9D9D9 ' Серый фон шапки, единый для таблиц Results.
Private Const FAILED_ROW_COLOR As Long = &HCEC7FF ' Светло-красный фон проблемного блока.
Private Const FAILED_FONT_COLOR As Long = &H9C0006 ' Темно-красный текст проблемного блока.
Private Const RESULTS_COLUMN_WIDTH As Double = 10# ' На Results оставляем одинаковую ширину столбцов.
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
    FormatTable target, batch
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
    target.Worksheet.Rows.Item(titleRow).Interior.Color = RGB(189, 215, 238)
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

' Очищает старую строку-название, чтобы при переносе якоря не оставались старые форматы.
Private Sub ClearBlockTitleRow(ByVal target As Object)
    If target.Row <= 1 Then Exit Sub
    target.Worksheet.Rows.Item(target.Row - 1).Clear
    target.Worksheet.Rows.Item(target.Row - 1).RowHeight = 15
End Sub

' Записывает многоуровневую шапку таблицы.
Private Sub WriteHeaders(ByVal target As Object, ByVal units As CUnitSystem)
    Dim headers() As Variant
    ReDim headers(1 To HEADER_ROWS, 1 To TOTAL_COLUMNS)

    headers(1, COL_ID) = "Combination ID"
    headers(1, COL_STATUS) = "общий статус"
    headers(1, COL_REF_N) = "относительно центральных осей приведенного сечения"
    headers(1, COL_LOAD_N) = "в точке приложения нагрузок"

    headers(1, COL_ARED) = "Момент образования трещин"
    headers(2, COL_ARED) = "Осевое растяжение"
    headers(2, COL_FORMATION_METHOD) = "Изгиб N + M"
    headers(2, COL_FORMATION_STATE) = "статус трещин"
    headers(1, COL_XT) = "нормальные и продольные трещины"
    headers(2, COL_XT) = "данные при изгибе N + M"
    headers(2, COL_DS_EQ) = "Расстояние между трещинами ls"
    headers(2, COL_PHI1) = "Коэффициенты"
    headers(2, COL_REBARS) = "Напряжения"
    headers(2, COL_ACRC) = "ширина раскрытия нормальных трещин"
    headers(2, COL_SIGMA_B_MAX) = "продольные трещины"

    PutHeaderNotes headers, units
    PutColumnLabels headers, units
    target.Resize(HEADER_ROWS, TOTAL_COLUMNS).Value2 = headers
    MergeHeaders target
End Sub

' Заполняет строку пояснений над нижним ярусом заголовков.
Private Sub PutHeaderNotes(ByRef headers() As Variant, ByVal units As CUnitSystem)
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
    headers(HEADER_NOTES_ROW, COL_REBARS) = "стержни, использованные для As и σs"
    headers(HEADER_NOTES_ROW, COL_NREBARS) = "количество выбранных стержней"
    headers(HEADER_NOTES_ROW, COL_SIGMA_S_CRC) = "напряжение выбранной арматуры сразу после образования трещины"
    headers(HEADER_NOTES_ROW, COL_SIGMA_S) = "напряжение выбранной арматуры при текущем НДС"
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
    headers(HEADER_LABEL_ROW, COL_FORMATION_STATE) = "state"
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
    MergeRange target, 1, COL_STATUS, HEADER_LABEL_ROW, COL_STATUS
    MergeRange target, 1, COL_REF_N, 2, COL_REF_MY
    MergeRange target, 1, COL_LOAD_N, 2, COL_LOAD_MY
    MergeRange target, 1, COL_ARED, 1, COL_FORMATION_STATE
    MergeRange target, 2, COL_ARED, 2, COL_LAMBDA_NCRC
    MergeRange target, 2, COL_FORMATION_METHOD, 2, COL_LAMBDA_MCRC
    MergeRange target, 2, COL_FORMATION_STATE, 3, COL_FORMATION_STATE
    MergeRange target, 1, COL_XT, 1, COL_MAIN_LAST
    MergeRange target, 2, COL_XT, 2, COL_AS
    MergeRange target, 2, COL_DS_EQ, 2, COL_LS
    MergeRange target, 2, COL_PHI1, 2, COL_PSI_S
    MergeRange target, 2, COL_REBARS, 2, COL_SIGMA_S
    MergeRange target, 2, COL_ACRC, 2, COL_ACRC_STATUS
    MergeRange target, 2, COL_SIGMA_B_MAX, 2, COL_LONG_STATUS
End Sub

' Записывает строки сочетаний одним массивом.
Private Sub WriteData(ByVal target As Object, ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    Dim dataRows As Long
    dataRows = target.Rows.Count - HEADER_ROWS
    If dataRows <= 0 Then Exit Sub

    Dim data() As Variant
    ReDim data(1 To dataRows, 1 To TOTAL_COLUMNS)

    Dim i As Long
    For i = 1 To dataRows
        If i <= batch.Count Then PutCombination data, i, batch, i, units
    Next i
    target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).Value2 = data
End Sub

' Заполняет одну строку сочетания фактически рассчитанными величинами трещин.
Private Sub PutCombination(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal units As CUnitSystem)
    Dim policy As CBatchStatusPolicy
    Set policy = StatusPolicy()

    data(rowIndex, COL_ID) = batch.CombinationID(index)
    data(rowIndex, COL_STATUS) = CrackOverallStatus(batch, index)

    If policy.DisplayStatus(batch.CrackStatus(index)) = policy.NotApplicable And _
            policy.DisplayStatus(batch.LongitudinalCrackStatus(index)) = policy.NotApplicable And _
            Not HasCrackFormationData(batch, index) Then
        Exit Sub
    End If

    data(rowIndex, COL_REF_N) = OutputForce(batch.CrackReferenceN(index), units)
    data(rowIndex, COL_REF_MX) = OutputMx(batch.CrackReferenceMx(index), units)
    data(rowIndex, COL_REF_MY) = OutputMy(batch.CrackReferenceMy(index), units)
    data(rowIndex, COL_LOAD_N) = OutputForce(batch.CrackLoadPointN(index), units)
    data(rowIndex, COL_LOAD_MX) = OutputMx(batch.CrackLoadPointMx(index), units)
    data(rowIndex, COL_LOAD_MY) = OutputMy(batch.CrackLoadPointMy(index), units)

    PutFormationData data, rowIndex, batch, index, units, policy
    PutCrackWidthData data, rowIndex, batch, index, units, policy
    PutLongitudinalData data, rowIndex, batch, index, units, policy
End Sub

' Записывает данные проверки стадии образования нормальной трещины.
Private Sub PutFormationData(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal units As CUnitSystem, ByVal policy As CBatchStatusPolicy)
    Dim hasFormationData As Boolean
    hasFormationData = HasCrackFormationData(batch, index)

    If policy.DisplayStatus(batch.CrackStatus(index)) = policy.NotApplicable And Not hasFormationData Then
        data(rowIndex, COL_FORMATION_STATE) = policy.NotApplicable
        Exit Sub
    End If

    If Len(batch.CrackFormationMethod(index)) > 0 Then _
        data(rowIndex, COL_FORMATION_METHOD) = batch.CrackFormationMethod(index)

    If batch.CrackNcrc(index) > SMALL_VALUE Then
        data(rowIndex, COL_ARED) = OutputArea(batch.CrackAred(index), units)
        data(rowIndex, COL_RBT_SER) = OutputStress(batch.CrackRbtSer(index), units)
        data(rowIndex, COL_NCRC) = OutputForce(batch.CrackNcrc(index), units)
        data(rowIndex, COL_LAMBDA_NCRC) = batch.CrackLambdaCrc(index)
    Else
        data(rowIndex, COL_NCRC) = policy.NotApplicable
    End If

    If (Not batch.CrackCentralBranch(index)) And hasFormationData Then
        data(rowIndex, COL_FORMATION_NCRC) = OutputForce(batch.CrackFormationNcrc(index), units)
    ElseIf Not batch.CrackCentralBranch(index) And batch.CrackMcrc(index) > SMALL_VALUE Then
        data(rowIndex, COL_FORMATION_NCRC) = policy.NotApplicable
    End If

    If batch.CrackMcrc(index) > SMALL_VALUE Then
        data(rowIndex, COL_MCRC) = OutputMomentMagnitude(batch.CrackMcrc(index), units)
        data(rowIndex, COL_LAMBDA_MCRC) = batch.CrackLambdaCrc(index)
    Else
        data(rowIndex, COL_MCRC) = policy.NotApplicable
    End If

    If policy.IsFinished(batch.CrackStatus(index)) Or hasFormationData Then
        data(rowIndex, COL_FORMATION_STATE) = IIf(batch.CrackFormed(index), "Cracked", "NotCracked")
    Else
        data(rowIndex, COL_FORMATION_STATE) = policy.DisplayStatus(batch.CrackStatus(index))
    End If
End Sub

' Записывает ширину раскрытия нормальных трещин и все величины ее формулы.
' Проверяет, есть ли сохраненные величины стадии образования нормальной трещины.
Private Function HasCrackFormationData(ByVal batch As CBatchSectionCalculator, ByVal index As Long) As Boolean
    HasCrackFormationData = batch.CrackNcrc(index) > SMALL_VALUE Or _
        Abs(batch.CrackFormationNcrc(index)) > SMALL_VALUE Or _
        batch.CrackMcrc(index) > SMALL_VALUE
End Function

' Записывает ширину раскрытия нормальных трещин и величины ее формулы.
Private Sub PutCrackWidthData(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal units As CUnitSystem, ByVal policy As CBatchStatusPolicy)
    If Not (policy.IsFinished(batch.CrackStatus(index)) And batch.CrackFormed(index)) Then
        If policy.DisplayStatus(batch.CrackStatus(index)) <> policy.NotApplicable Then _
            data(rowIndex, COL_ACRC_STATUS) = policy.DisplayStatus(batch.CrackStatus(index))
        Exit Sub
    End If

    data(rowIndex, COL_XT) = OutputLength(batch.CrackTensionDepth(index), units)
    data(rowIndex, COL_H) = OutputLength(batch.CrackSectionDepthH(index), units)
    data(rowIndex, COL_AS_COVER) = OutputLength(batch.CrackCoverA(index), units)
    data(rowIndex, COL_HBT) = OutputLength(batch.CrackEffectiveZoneDepth(index), units)
    data(rowIndex, COL_ABT) = OutputArea(batch.CrackAbt(index), units)
    data(rowIndex, COL_AS) = OutputArea(batch.CrackAs(index), units)
    data(rowIndex, COL_DS_EQ) = OutputLength(batch.CrackDsEquivalent(index), units)
    data(rowIndex, COL_LS0) = OutputLength(batch.CrackSpacingRaw(index), units)
    data(rowIndex, COL_LS) = OutputLength(batch.CrackSpacing(index), units)
    data(rowIndex, COL_PHI1) = batch.CrackPhi1(index)
    data(rowIndex, COL_PHI2) = batch.CrackPhi2(index)
    data(rowIndex, COL_PHI3) = batch.CrackPhi3(index)
    data(rowIndex, COL_PSI_S) = batch.CrackPsiS(index)
    data(rowIndex, COL_REBARS) = batch.CrackTensionRebarIds(index)
    data(rowIndex, COL_NREBARS) = batch.CrackTensionRebarCount(index)
    data(rowIndex, COL_SIGMA_S_CRC) = OptionalStress(batch.CrackSigmaSCrc(index), units, policy)
    data(rowIndex, COL_SIGMA_S) = OutputStress(batch.CrackSigmaS(index), units)
    data(rowIndex, COL_ACRC) = OutputLength(batch.CrackWidth(index), units)
    data(rowIndex, COL_ACRC_ULT) = OutputLength(batch.CrackAllowable(index), units)
    data(rowIndex, COL_ACRC_RATIO) = CrackReserveFactor(batch, index)
    data(rowIndex, COL_ACRC_STATUS) = policy.DisplayStatus(batch.CrackStatus(index))
End Sub

' Записывает проверку продольных трещин по сжатому бетону.
Private Sub PutLongitudinalData(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal units As CUnitSystem, ByVal policy As CBatchStatusPolicy)
    If Not policy.IsFinished(batch.LongitudinalCrackStatus(index)) Then
        data(rowIndex, COL_LONG_STATUS) = policy.DisplayStatus(batch.LongitudinalCrackStatus(index))
        Exit Sub
    End If

    data(rowIndex, COL_SIGMA_B_MAX) = OutputStress(batch.MaxConcreteCompressionStress(index), units)
    data(rowIndex, COL_RB_MC2) = OutputStress(batch.LongitudinalCrackRbMc2(index), units)
    data(rowIndex, COL_LONG_RATIO) = LongitudinalCrackReserveFactor(batch, index)
    data(rowIndex, COL_LONG_STATUS) = policy.DisplayStatus(batch.LongitudinalCrackStatus(index))
End Sub

' Возвращает запас по нормальным трещинам для подробной пользовательской таблицы:
' допустимая ширина делится на расчетную. При нулевой ширине запас бесконечный,
' поэтому ячейка остается пустой, как и в верхнем batch summary.
Private Function CrackReserveFactor(ByVal batch As CBatchSectionCalculator, ByVal index As Long) As Variant
    If batch.CrackUtilization(index) > SMALL_VALUE Then _
        CrackReserveFactor = 1# / batch.CrackUtilization(index)
End Function

' Возвращает запас по продольным трещинам: Rb,mc2 делится на фактическое
' сжатие бетона. При отсутствии сжатия ячейка остается пустой.
Private Function LongitudinalCrackReserveFactor(ByVal batch As CBatchSectionCalculator, ByVal index As Long) As Variant
    If batch.LongitudinalCrackUtilization(index) > SMALL_VALUE Then _
        LongitudinalCrackReserveFactor = 1# / batch.LongitudinalCrackUtilization(index)
End Function

' Оформляет шапки, границы, числовые форматы и подсветку статусов.
Private Sub FormatTable(ByVal target As Object, ByVal batch As CBatchSectionCalculator)
    target.Font.Name = "Arial"
    target.Font.Size = 9
    target.HorizontalAlignment = -4108
    target.VerticalAlignment = -4108
    target.WrapText = False
    target.Rows.RowHeight = 15

    SetColumnWidths target
    FormatBlock target, COL_LEFT_FIRST, COL_LEFT_LAST
    FormatBlock target, COL_MAIN_FIRST, COL_MAIN_LAST
    target.Columns(COL_SEPARATOR).Clear

    target.Resize(HEADER_ROWS, TOTAL_COLUMNS).HorizontalAlignment = -4108
    target.Resize(HEADER_ROWS, TOTAL_COLUMNS).VerticalAlignment = -4108
    target.Resize(HEADER_ROWS, TOTAL_COLUMNS).WrapText = True
    target.Rows(HEADER_NOTES_ROW).Font.Bold = False
    target.Rows(HEADER_NOTES_ROW).HorizontalAlignment = -4131
    target.Rows(HEADER_NOTES_ROW).WrapText = False
    target.Rows(HEADER_LABEL_ROW).WrapText = False
    target.Cells(HEADER_LABEL_ROW, COL_RBT_SER).WrapText = False

    Dim dataRows As Long
    dataRows = target.Rows.Count - HEADER_ROWS
    If dataRows > 0 Then
        target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).HorizontalAlignment = -4108
        target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).VerticalAlignment = -4108
        target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).NumberFormat = "0.000"
        target.Offset(HEADER_ROWS, COL_ID - 1).Resize(dataRows, 2).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_REBARS - 1).Resize(dataRows, 1).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_FORMATION_STATE - 1).Resize(dataRows, 1).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_ACRC_STATUS - 1).Resize(dataRows, 1).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_LONG_STATUS - 1).Resize(dataRows, 1).NumberFormat = "@"
        ApplyFailedRows target, batch, dataRows
    End If
    ApplyHeaderSubscripts target
End Sub

' Задает одинаковую стандартную ширину всем столбцам блока Results.
Private Sub SetColumnWidths(ByVal target As Object)
    Dim columnIndex As Long
    For columnIndex = 1 To TOTAL_COLUMNS
        target.Worksheet.Columns(target.Column + columnIndex - 1).ColumnWidth = RESULTS_COLUMN_WIDTH
    Next columnIndex
End Sub

' Оформляет один непрерывный блок таблицы, не затрагивая пустые разделители.
Private Sub FormatBlock(ByVal target As Object, ByVal firstCol As Long, ByVal lastCol As Long)
    Dim block As Object
    Set block = target.Offset(0, firstCol - 1).Resize(target.Rows.Count, lastCol - firstCol + 1)
    block.Borders.LineStyle = 1
    block.Borders.Weight = 2
    block.Rows(1).Resize(HEADER_ROWS).Font.Bold = True
    target.Offset(0, firstCol - 1).Resize(HEADER_ROWS, lastCol - firstCol + 1).Interior.Color = HEADER_COLOR
End Sub

' Подсвечивает только те части таблицы, где проверка трещин не прошла.
Private Sub ApplyFailedRows(ByVal target As Object, ByVal batch As CBatchSectionCalculator, ByVal dataRows As Long)
    Dim policy As CBatchStatusPolicy
    Set policy = StatusPolicy()

    Dim i As Long
    For i = 1 To dataRows
        If i <= batch.Count Then
            If policy.IsFailedOverall(batch.CrackStatus(i)) Or policy.IsFailedOverall(batch.LongitudinalCrackStatus(i)) Then
                PaintFailed target.Offset(HEADER_ROWS + i - 1, 0).Resize(1, COL_LEFT_LAST)
            End If
            If policy.IsFailedOverall(batch.CrackStatus(i)) Then
                PaintFailed target.Offset(HEADER_ROWS + i - 1, COL_MAIN_FIRST - 1).Resize(1, COL_ACRC_STATUS - COL_MAIN_FIRST + 1)
            End If
            If policy.IsFailedOverall(batch.LongitudinalCrackStatus(i)) Then
                PaintFailed target.Offset(HEADER_ROWS + i - 1, COL_SIGMA_B_MAX - 1).Resize(1, COL_LONG_STATUS - COL_SIGMA_B_MAX + 1)
            End If
        End If
    Next i
End Sub

' Красит проблемный фрагмент строки.
Private Sub PaintFailed(ByVal rangeObject As Object)
    With rangeObject
        .Interior.Color = FAILED_ROW_COLOR
        .Font.Color = FAILED_FONT_COLOR
    End With
End Sub

' Выставляет нижние индексы в обозначениях нижнего яруса заголовков.
Private Sub ApplyHeaderSubscripts(ByVal target As Object)
    Dim labels As Object
    Set labels = target.Rows(HEADER_LABEL_ROW)

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
    startPos = InStr(1, textValue, token, vbTextCompare)
    Do While startPos > 0
        cell.Characters(startPos, Len(token)).Font.Subscript = True
        startPos = InStr(startPos + Len(token), textValue, token, vbTextCompare)
    Loop
End Sub

' Возвращает общий статус именно для crack-блока.
Private Function CrackOverallStatus(ByVal batch As CBatchSectionCalculator, ByVal index As Long) As String
    CrackOverallStatus = StatusPolicy().Aggregate( _
        StatusPolicy().NotApplicable, StatusPolicy().NotApplicable, _
        batch.CrackStatus(index), batch.LongitudinalCrackStatus(index))
End Function

' Возвращает напряжение или N/A для величин, которые расчетно могут отсутствовать.
Private Function OptionalStress(ByVal value As Double, ByVal units As CUnitSystem, ByVal policy As CBatchStatusPolicy) As Variant
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
Private Function StatusPolicy() As CBatchStatusPolicy
    Static policy As CBatchStatusPolicy
    If policy Is Nothing Then Set policy = New CBatchStatusPolicy
    Set StatusPolicy = policy
End Function
